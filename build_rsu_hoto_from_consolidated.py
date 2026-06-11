from __future__ import annotations

from datetime import datetime
from pathlib import Path
import argparse
import sys

import pandas as pd

from build_rsu_master_from_consolidated import (
    FIBER_ITEM_PREFIX,
    FIBER_ORDER,
    _excel_serial_to_datetime,
    load_cluster_to_rsu_mapping,
    normalize_fiber_capacity,
    normalize_link_type_to_tpt_status,
    read_excel_lock_safe,
    resolve_cluster_to_rsu,
    to_nullable_int,
)


OUTPUT_COLUMNS = [
    "ID",
    "RSU Code",
    "Hoto Date",
    "TPT Status",
    "ID (Delivered RSU)",
    "Delivered FAT (Delivered RSU)",
    "Delivered Fibre (Delivered RSU)",
    "Delivered HP (Delivered RSU)",
    "Hoto Item (Delivered RSU)",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert consolidated_hp.xlsx into an ERPNext-importable 'RSU Hoto' xlsx "
            "(one parent row per RSU Code + HOTO Date + Link Type + 4 fiber capacity child rows)."
        )
    )
    parser.add_argument("--input-file", default="consolidated_hp.xlsx")
    parser.add_argument("--input-sheet", default="Consolidated")
    parser.add_argument("--output-file", default="RSU Hoto Generated.xlsx")
    parser.add_argument("--output-sheet", default="RSU Hoto")
    parser.add_argument(
        "--cluster-site-file",
        default="cluster_site_comparison.xlsx",
        help=(
            'Excel file with columns "Cluster ID" and "total match". '
            "RSU Code is taken from total match for Cluster IDs present in "
            "the consolidated input."
        ),
    )
    parser.add_argument(
        "--cluster-site-sheet",
        default=None,
        help="Sheet name in the cluster/site comparison file (default: first sheet).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input_file).resolve()
    output_path = Path(args.output_file).resolve()

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        return 1

    df = read_excel_lock_safe(input_path, args.input_sheet)

    comparison_path = Path(args.cluster_site_file).resolve()
    cluster_to_rsu = load_cluster_to_rsu_mapping(comparison_path, args.cluster_site_sheet)
    if not cluster_to_rsu:
        print("No Cluster ID -> RSU Code mapping loaded. Cannot build RSU Hoto.")
        return 1

    print(
        f"Loaded {len(cluster_to_rsu)} Cluster ID -> RSU Code entries "
        f"from {comparison_path.name}."
    )

    required = [
        "Cluster ID",
        "Fiber Capacity",
        "Fiber Length as per HOTO",
        "HOTO Date",
        "Actual Deployed FAT",
        "Actual Home Pass",
    ]
    missing = [column for column in required if column not in df.columns]
    if missing:
        print(f"Missing required columns in input: {missing}")
        return 1

    df = df.copy()
    df = df.dropna(subset=["Cluster ID"])
    df["Cluster ID"] = df["Cluster ID"].astype(str).str.strip()
    df = df[(df["Cluster ID"] != "") & (df["Cluster ID"].str.lower() != "nan")]

    df["RSU Code"] = df["Cluster ID"].map(
        lambda cluster_id: resolve_cluster_to_rsu(cluster_id, cluster_to_rsu)
    )
    unmatched_mask = df["RSU Code"].isna()
    unmatched_clusters = sorted(df.loc[unmatched_mask, "Cluster ID"].unique().tolist())
    df = df[~unmatched_mask].copy()
    if unmatched_clusters:
        print(
            f"Skipped {len(unmatched_clusters)} Cluster ID(s) with no RSU Code in "
            f"{comparison_path.name} (after normalized comparison)."
        )
        for cluster_id in unmatched_clusters:
            print(f"  {cluster_id!r}")

    if df.empty:
        print("No rows left after matching Cluster IDs to RSU Codes.")
        return 1

    if "Link Type" in df.columns:
        df["TPT Status"] = df["Link Type"].map(normalize_link_type_to_tpt_status)
    else:
        df["TPT Status"] = ""

    df["Fiber Capacity"] = df["Fiber Capacity"].map(normalize_fiber_capacity)
    df = df[df["Fiber Capacity"].isin(FIBER_ORDER)]

    df["Fiber Length as per HOTO"] = pd.to_numeric(
        df["Fiber Length as per HOTO"], errors="coerce"
    ).fillna(0)
    df["Actual Deployed FAT"] = pd.to_numeric(df["Actual Deployed FAT"], errors="coerce")
    df["Actual Home Pass"] = pd.to_numeric(df["Actual Home Pass"], errors="coerce")
    df["HOTO Date"] = pd.to_datetime(
        df["HOTO Date"].map(_excel_serial_to_datetime), errors="coerce"
    )

    rows_missing_date = int(df["HOTO Date"].isna().sum())
    df = df.dropna(subset=["HOTO Date"])

    df = df.sort_values(
        by=[
            "RSU Code",
            "HOTO Date",
            "TPT Status",
            "Fiber Capacity",
            "Fiber Length as per HOTO",
        ],
        ascending=[True, True, True, True, False],
    )
    df = df.drop_duplicates(
        subset=["RSU Code", "HOTO Date", "TPT Status", "Fiber Capacity"],
        keep="first",
    )

    rows: list[dict[str, object]] = []

    for (rsu_code, hoto_date_value, tpt_status_value), group in df.groupby(
        ["RSU Code", "HOTO Date", "TPT Status"], sort=True
    ):
        group_by_capacity = group.set_index("Fiber Capacity")

        fat_series = group["Actual Deployed FAT"].dropna()
        fat_value = fat_series.max() if not fat_series.empty else ""
        hp_series = group["Actual Home Pass"].dropna()
        hp_value = hp_series.max() if not hp_series.empty else ""

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
                row["Hoto Date"] = hoto_date_value
                row["TPT Status"] = tpt_status_value
            row["Delivered FAT (Delivered RSU)"] = fat_value
            row["Delivered Fibre (Delivered RSU)"] = fiber_length_value
            row["Delivered HP (Delivered RSU)"] = hp_value
            row["Hoto Item (Delivered RSU)"] = item_code
            rows.append(row)

    out_df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    numeric_columns = [
        "Delivered FAT (Delivered RSU)",
        "Delivered Fibre (Delivered RSU)",
        "Delivered HP (Delivered RSU)",
    ]
    for column in numeric_columns:
        out_df[column] = out_df[column].map(to_nullable_int).astype("Int64")

    out_df["Hoto Date"] = pd.to_datetime(out_df["Hoto Date"], errors="coerce").dt.date

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, sheet_name=args.output_sheet, index=False)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = output_path.with_name(
            f"{output_path.stem}_{timestamp}{output_path.suffix}"
        )
        print(
            f"'{output_path.name}' is open or write-locked. "
            f"Writing to '{fallback_path.name}' instead."
        )
        with pd.ExcelWriter(fallback_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, sheet_name=args.output_sheet, index=False)
        output_path = fallback_path

    group_count = len(out_df) // len(FIBER_ORDER)
    distinct_rsus = int(
        out_df["RSU Code"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    )
    print(f"Output: {output_path}")
    print(f"(RSU Code x HOTO Date x Link Type) groups written: {group_count}")
    print(f"Distinct RSU Codes: {distinct_rsus}")
    print(f"Total rows (groups x {len(FIBER_ORDER)}): {len(out_df)}")
    if rows_missing_date:
        print(
            f"Skipped {rows_missing_date} source row(s) with no HOTO Date "
            "(cannot be assigned to a date-specific block)."
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
