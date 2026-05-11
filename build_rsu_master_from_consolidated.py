from __future__ import annotations

from datetime import datetime
from pathlib import Path
import argparse
import io
import os
import re
import sys

import pandas as pd


# Map full circle names found in consolidated_hp.xlsx to ERPNext circle codes
# used in the existing "RSU Master.csv".
CIRCLE_TO_CODE: dict[str, str] = {
    "Chandigarh": "PUN",
    "Punjab": "PUN",
    "Delhi": "DEL",
    "Karnataka": "KTK",
    "Bangalore": "KTK",
    "Kolkata": "KOL",
    "Mumbai": "M&G",
    "M&G": "M&G",
    "Pune": "M&G",
    "Tamil Nadu": "TAM",
    "UPW": "UPW",
}

FIBER_ITEM_PREFIX = "SER-FSDU-CFT-"
FIBER_ORDER = ["6F", "12F", "24F", "48F"]

OUTPUT_COLUMNS = [
    "ID",
    "RSU Code",
    "SDU Billing Rates",
    "Circle",
    "Decom Date",
    "ID (Rsu Items)",
    "Item (Rsu Items)",
    "Plan FAT (Rsu Items)",
    "Plan Fibre length (Rsu Items)",
    "Plan HP (Rsu Items)",
]


def normalize_fiber_capacity(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper().replace(" ", "")
    if not text:
        return ""
    if not text.endswith("F"):
        text = text + "F"
    return text


def map_circle_to_code(circle_text: str) -> str:
    text = circle_text.strip()
    return CIRCLE_TO_CODE.get(text, text)


def to_nullable_int(value: object) -> object:
    if value is None or value == "":
        return pd.NA
    try:
        if pd.isna(value):
            return pd.NA
    except (TypeError, ValueError):
        pass
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return pd.NA


def _read_bytes_with_shared_access_windows(path: Path) -> bytes:
    """Read a file's bytes on Windows using CreateFileW with full sharing
    flags so we can read even when Excel holds an exclusive write lock."""
    import ctypes
    from ctypes import wintypes

    GENERIC_READ = 0x80000000
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
    FILE_ATTRIBUTE_NORMAL = 0x80

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    kernel32.CreateFileW.restype = wintypes.HANDLE
    kernel32.CreateFileW.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
        ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
    ]
    kernel32.GetFileSizeEx.restype = wintypes.BOOL
    kernel32.GetFileSizeEx.argtypes = [wintypes.HANDLE, ctypes.POINTER(ctypes.c_longlong)]
    kernel32.ReadFile.restype = wintypes.BOOL
    kernel32.ReadFile.argtypes = [
        wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD), ctypes.c_void_p,
    ]
    kernel32.CloseHandle.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    handle = kernel32.CreateFileW(
        str(path),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None,
    )
    if not handle or handle == INVALID_HANDLE_VALUE:
        raise OSError(ctypes.get_last_error(), f"CreateFileW failed for {path}")

    try:
        file_size = ctypes.c_longlong(0)
        if not kernel32.GetFileSizeEx(handle, ctypes.byref(file_size)):
            raise OSError(ctypes.get_last_error(), "GetFileSizeEx failed")

        remaining = file_size.value
        chunks: list[bytes] = []
        chunk_size = 1024 * 1024
        buffer = ctypes.create_string_buffer(chunk_size)
        bytes_read = wintypes.DWORD(0)

        while remaining > 0:
            to_read = min(chunk_size, remaining)
            ok = kernel32.ReadFile(
                handle, buffer, to_read, ctypes.byref(bytes_read), None
            )
            if not ok:
                raise OSError(ctypes.get_last_error(), "ReadFile failed")
            if bytes_read.value == 0:
                break
            chunks.append(buffer.raw[: bytes_read.value])
            remaining -= bytes_read.value

        return b"".join(chunks)
    finally:
        kernel32.CloseHandle(handle)


def read_excel_lock_safe(input_path: Path, sheet_name: str) -> pd.DataFrame:
    """Read an Excel sheet. Falls back to a Windows shared-read on
    PermissionError so the script works while the file is open in Excel."""
    try:
        return pd.read_excel(input_path, sheet_name=sheet_name)
    except PermissionError:
        if os.name != "nt":
            raise
        print(
            f"'{input_path.name}' appears to be open (locked). "
            "Reading via a Windows shared-read."
        )
        data = _read_bytes_with_shared_access_windows(input_path)
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert consolidated_hp.xlsx into an ERPNext-importable RSU Master CSV "
            "(parent RSU row + 4 fiber capacity child rows: 6F / 12F / 24F / 48F)."
        )
    )
    parser.add_argument("--input-file", default="consolidated_hp.xlsx")
    parser.add_argument("--input-sheet", default="Consolidated")
    parser.add_argument("--output-file", default="RSU Master Generated.csv")
    parser.add_argument(
        "--rsu-suffix",
        default="-0",
        help='Suffix appended to Cluster ID to form RSU Code (default: "-0", matching the existing RSU Master template). Pass "" for no suffix.',
    )
    parser.add_argument(
        "--mapping-file",
        default="mapping.xlsx",
        help=(
            'Optional Excel file with columns "Cluster ID" and "Site id". '
            "When a Cluster ID matches, the RSU Code is taken verbatim from "
            "Site id (the --rsu-suffix is not appended). Skipped silently if "
            "the file is missing."
        ),
    )
    parser.add_argument(
        "--mapping-sheet",
        default=None,
        help="Sheet name in the mapping file (default: first sheet).",
    )
    return parser.parse_args()


def load_cluster_to_site_mapping(
    mapping_path: Path, sheet_name: str | None
) -> dict[str, str]:
    """Read mapping Excel and return {Cluster ID -> Site id}.
    Both keys and values are stripped of surrounding whitespace. Empty rows
    are skipped. Returns an empty dict if the file does not exist."""
    if not mapping_path.exists():
        return {}

    effective_sheet: object = sheet_name if sheet_name is not None else 0
    df = read_excel_lock_safe(mapping_path, effective_sheet)

    cluster_col = next(
        (c for c in df.columns if str(c).strip().lower() == "cluster id"), None
    )
    site_col = next(
        (c for c in df.columns if str(c).strip().lower() == "site id"), None
    )
    if cluster_col is None or site_col is None:
        print(
            f"Mapping file '{mapping_path.name}' is missing required columns "
            "'Cluster ID' and/or 'Site id'. Skipping mapping."
        )
        return {}

    mapping: dict[str, str] = {}
    for cluster_value, site_value in zip(df[cluster_col], df[site_col]):
        if pd.isna(cluster_value) or pd.isna(site_value):
            continue
        cluster_key = str(cluster_value).strip()
        site_id = str(site_value).strip()
        if cluster_key and site_id:
            mapping[cluster_key] = site_id
    return mapping


def main() -> int:
    args = parse_args()
    input_path = Path(args.input_file).resolve()
    output_path = Path(args.output_file).resolve()

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        return 1

    df = read_excel_lock_safe(input_path, args.input_sheet)

    mapping_path = Path(args.mapping_file).resolve()
    cluster_to_site = load_cluster_to_site_mapping(mapping_path, args.mapping_sheet)
    if cluster_to_site:
        print(
            f"Loaded {len(cluster_to_site)} Cluster ID -> Site id overrides "
            f"from {mapping_path.name}."
        )

    required = [
        "Cluster ID",
        "Fiber Capacity",
        "Planned FAT",
        "Planned HP",
        "Fiber Length as per HOTO",
        "Circle",
    ]
    missing = [column for column in required if column not in df.columns]
    if missing:
        print(f"Missing required columns in input: {missing}")
        return 1

    df = df.copy()
    # Drop NaN cluster IDs first; otherwise pandas' string dtype keeps them
    # as <NA> which slips past textual filters and produces "NAN-0".
    df = df.dropna(subset=["Cluster ID"])
    df["Cluster ID"] = df["Cluster ID"].astype(str).str.strip()
    df = df[(df["Cluster ID"] != "") & (df["Cluster ID"].str.lower() != "nan")]

    df["Fiber Capacity"] = df["Fiber Capacity"].map(normalize_fiber_capacity)
    df = df[df["Fiber Capacity"].isin(FIBER_ORDER)]

    df["Fiber Length as per HOTO"] = pd.to_numeric(
        df["Fiber Length as per HOTO"], errors="coerce"
    ).fillna(0)
    df["Planned FAT"] = pd.to_numeric(df["Planned FAT"], errors="coerce")
    df["Planned HP"] = pd.to_numeric(df["Planned HP"], errors="coerce")

    # Resolve each Cluster ID to its final RSU Code BEFORE dedup/groupby.
    # When the mapping has a match, use the mapped Site id verbatim (so e.g.
    # "1JL - TPT" and "1JL" both collapse to "1JL-0"). Otherwise apply the
    # default <Cluster ID><suffix> rule. Doing this up front means clusters
    # that resolve to the same RSU Code are merged by the existing dedup
    # logic instead of producing duplicate parent rows.
    mapped_cluster_ids: set[str] = set()
    multi_value_mappings: list[tuple[str, str]] = []

    def resolve_rsu_code(cluster_id: str) -> str:
        site_id = cluster_to_site.get(cluster_id)
        if site_id:
            mapped_cluster_ids.add(cluster_id)
            # Some mapping cells contain multiple Site ids joined with "&"
            # (e.g. "FO9-0 & FO9-1"). ERPNext requires one value per cell so
            # we take the first one and warn afterwards.
            if "&" in site_id:
                multi_value_mappings.append((cluster_id, site_id))
                site_id = site_id.split("&", 1)[0]
            raw_code = site_id
        else:
            raw_code = f"{cluster_id}{args.rsu_suffix}"
        # Strip ALL internal whitespace and uppercase so codes like
        # "3BZ -0", "33T - 1-0", "GULC - TPT-0" become "3BZ-0", "33T-1-0",
        # "GULC-TPT-0" -- matching the Site doctype values in ERPNext.
        return re.sub(r"\s+", "", raw_code).upper()

    df["RSU Code"] = df["Cluster ID"].map(resolve_rsu_code)

    # Count, on the pre-dedup view, how many target RSU Codes have 2+ distinct
    # source Cluster IDs feeding into them (i.e., real merges).
    sources_per_rsu = df.groupby("RSU Code")["Cluster ID"].nunique()
    merged_groups = int((sources_per_rsu >= 2).sum())

    # When the same (RSU Code, Fiber Capacity) appears multiple times across
    # files / HOTO revisions / merged clusters, prefer the row with the
    # largest measured fiber length so blanks/zeros don't overwrite real
    # measurements.
    df = df.sort_values(
        by=["RSU Code", "Fiber Capacity", "Fiber Length as per HOTO"],
        ascending=[True, True, False],
    )
    df = df.drop_duplicates(subset=["RSU Code", "Fiber Capacity"], keep="first")

    rows: list[dict[str, object]] = []

    for rsu_code, group in df.groupby("RSU Code", sort=True):
        group_by_capacity = group.set_index("Fiber Capacity")

        circle_values = (
            group["Circle"].dropna().astype(str).str.strip()
        )
        circle_values = circle_values[circle_values != ""]
        circle_text = circle_values.mode().iloc[0] if not circle_values.empty else ""
        circle_code = map_circle_to_code(circle_text)
        billing_rate = f"{circle_code} 1 Year Fix" if circle_code else ""

        plan_fat_series = group["Planned FAT"].dropna()
        plan_fat_value = plan_fat_series.max() if not plan_fat_series.empty else ""
        plan_hp_series = group["Planned HP"].dropna()
        plan_hp_value = plan_hp_series.max() if not plan_hp_series.empty else ""

        for index, capacity in enumerate(FIBER_ORDER):
            item_code = f"{FIBER_ITEM_PREFIX}{capacity}"

            if capacity in group_by_capacity.index:
                fiber_length_value = group_by_capacity.loc[capacity, "Fiber Length as per HOTO"]
                if isinstance(fiber_length_value, pd.Series):
                    fiber_length_value = fiber_length_value.iloc[0]
            else:
                fiber_length_value = 0

            row: dict[str, object] = {column: "" for column in OUTPUT_COLUMNS}
            if index == 0:
                row["RSU Code"] = rsu_code
                row["SDU Billing Rates"] = billing_rate
                row["Circle"] = circle_code
            row["Item (Rsu Items)"] = item_code
            row["Plan FAT (Rsu Items)"] = plan_fat_value
            row["Plan Fibre length (Rsu Items)"] = fiber_length_value
            row["Plan HP (Rsu Items)"] = plan_hp_value
            rows.append(row)

    out_df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    numeric_columns = [
        "Plan FAT (Rsu Items)",
        "Plan Fibre length (Rsu Items)",
        "Plan HP (Rsu Items)",
    ]
    for column in numeric_columns:
        out_df[column] = out_df[column].map(to_nullable_int).astype("Int64")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        out_df.to_csv(output_path, index=False)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = output_path.with_name(
            f"{output_path.stem}_{timestamp}{output_path.suffix}"
        )
        print(
            f"'{output_path.name}' is open or write-locked. "
            f"Writing to '{fallback_path.name}' instead."
        )
        out_df.to_csv(fallback_path, index=False)
        output_path = fallback_path

    rsu_count = len(out_df) // len(FIBER_ORDER)
    print(f"Output: {output_path}")
    print(f"RSUs written: {rsu_count}")
    print(f"Total rows (RSUs x {len(FIBER_ORDER)}): {len(out_df)}")
    if cluster_to_site:
        print(
            f"Cluster IDs matched by mapping file: {len(mapped_cluster_ids)} "
            f"(of {len(cluster_to_site)} mapping entries)."
        )
        print(f"RSU Codes that merged 2+ source Cluster IDs: {merged_groups}")

    if multi_value_mappings:
        print(
            f"\nWARNING: {len(multi_value_mappings)} mapping value(s) contain "
            "'&' (multiple Site ids in one cell). The first value was used; "
            "fix mapping.xlsx if a different choice is needed:"
        )
        for cluster_id, site_value in multi_value_mappings:
            print(f"  {cluster_id!r} -> {site_value!r}")

    parent_codes = (
        out_df["RSU Code"].dropna().astype(str).str.strip()
    )
    parent_codes = parent_codes[parent_codes != ""]
    suspicious = sorted(
        set(parent_codes[~parent_codes.str.fullmatch(r"[A-Z0-9]+-\d+")].tolist())
    )
    if suspicious:
        print(
            f"\nWARNING: {len(suspicious)} RSU Code(s) do not match the "
            "expected '<ID>-<digit>' pattern. ERPNext may reject these. "
            "Consider adding/fixing entries in mapping.xlsx:"
        )
        for code in suspicious:
            print(f"  {code}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
