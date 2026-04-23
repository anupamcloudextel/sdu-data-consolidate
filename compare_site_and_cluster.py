from __future__ import annotations

from pathlib import Path
import argparse
import sys

import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compare 'Site ID' from Site.xlsx with 'Cluster ID' from "
            "consolidated_hp.xlsx after appending '-0' to Cluster ID values."
        )
    )
    parser.add_argument(
        "--site-file",
        default="Site.xlsx",
        help="Path to Site.xlsx file (default: Site.xlsx).",
    )
    parser.add_argument(
        "--site-sheet",
        default=0,
        help="Sheet name/index in Site.xlsx (default: first sheet).",
    )
    parser.add_argument(
        "--cluster-file",
        default="consolidated_hp.xlsx",
        help="Path to consolidated_hp.xlsx file (default: consolidated_hp.xlsx).",
    )
    parser.add_argument(
        "--cluster-sheet",
        default="Consolidated",
        help="Sheet name in consolidated_hp.xlsx (default: Consolidated).",
    )
    parser.add_argument(
        "--output-file",
        default="site_cluster_comparison.xlsx",
        help="Output Excel filename (default: site_cluster_comparison.xlsx).",
    )
    return parser.parse_args()


def to_clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_id_value(value: object) -> str:
    return "".join(to_clean_text(value).split()).upper()


def normalize_column_name(column_name: object) -> str:
    return "".join(str(column_name).strip().lower().split())


def get_column_by_flexible_name(dataframe: pd.DataFrame, expected_name: str) -> str | None:
    target = normalize_column_name(expected_name)
    for column in dataframe.columns:
        if normalize_column_name(column) == target:
            return str(column)
    return None


def main() -> int:
    args = parse_args()

    site_file = Path(args.site_file).resolve()
    cluster_file = Path(args.cluster_file).resolve()
    output_file = Path(args.output_file).resolve()

    if not site_file.exists():
        print(f"Site file not found: {site_file}")
        return 1
    if not cluster_file.exists():
        print(f"Cluster file not found: {cluster_file}")
        return 1

    try:
        site_df = pd.read_excel(site_file, sheet_name=args.site_sheet)
    except Exception as exc:
        print(f"Failed to read site file: {site_file} -> {exc}")
        return 1

    try:
        cluster_df = pd.read_excel(cluster_file, sheet_name=args.cluster_sheet)
    except Exception as exc:
        print(f"Failed to read cluster file: {cluster_file} -> {exc}")
        return 1

    site_id_col = get_column_by_flexible_name(site_df, "Site ID")
    cluster_id_col = get_column_by_flexible_name(cluster_df, "Cluster ID")
    circle_col = get_column_by_flexible_name(cluster_df, "Circle")

    if site_id_col is None:
        print("Column 'Site ID' not found in site file.")
        print("Available columns in site file:")
        print(", ".join(map(str, site_df.columns)))
        return 1
    if cluster_id_col is None:
        print("Column 'Cluster ID' not found in cluster file.")
        print("Available columns in cluster file:")
        print(", ".join(map(str, cluster_df.columns)))
        return 1

    site_id_lookup: dict[str, str] = {}
    for value in site_df[site_id_col].tolist():
        original = to_clean_text(value)
        normalized = normalize_id_value(original)
        if normalized and normalized not in site_id_lookup:
            site_id_lookup[normalized] = original

    site_id_set = set(site_id_lookup.keys())

    working_df = cluster_df.copy()
    working_df["Cluster ID"] = working_df[cluster_id_col].map(to_clean_text)
    working_df = working_df[working_df["Cluster ID"] != ""].copy()
    if circle_col is not None:
        working_df["Circle"] = working_df[circle_col].map(to_clean_text)
    else:
        working_df["Circle"] = ""

    grouped = (
        working_df.groupby("Cluster ID", dropna=False)
        .agg(
            **{
                "Count in consolidated_hp": ("Cluster ID", "size"),
                "Circle": (
                    "Circle",
                    lambda series: ", ".join(
                        value for value in dict.fromkeys(series.tolist()) if value
                    ),
                ),
            }
        )
        .reset_index()
    )
    grouped["Cluster ID + -0"] = grouped["Cluster ID"].map(lambda x: f"{x}-0")
    grouped["Cluster ID + -1"] = grouped["Cluster ID"].map(lambda x: f"{x}-1")
    grouped["Compare Key (direct)"] = grouped["Cluster ID"].map(normalize_id_value)
    grouped["Compare Key -0"] = grouped["Cluster ID + -0"].map(normalize_id_value)
    grouped["Compare Key -1"] = grouped["Cluster ID + -1"].map(normalize_id_value)
    grouped["Site ID (direct)"] = grouped["Cluster ID"].map(
        lambda value: site_id_lookup.get(normalize_id_value(value), "")
    )
    grouped["Site ID (-0)"] = grouped["Cluster ID + -0"].map(
        lambda value: site_id_lookup.get(normalize_id_value(value), "")
    )
    grouped["Site ID (-1)"] = grouped["Cluster ID + -1"].map(
        lambda value: site_id_lookup.get(normalize_id_value(value), "")
    )
    grouped["Match (direct)"] = grouped["Compare Key (direct)"].map(
        lambda value: "TRUE" if value in site_id_set else "FALSE"
    )
    grouped["Match -0"] = grouped["Compare Key -0"].map(
        lambda value: "TRUE" if value in site_id_set else "FALSE"
    )
    grouped["Match -1"] = grouped["Compare Key -1"].map(
        lambda value: "TRUE" if value in site_id_set else "FALSE"
    )
    grouped["Site ID"] = grouped.apply(
        lambda row: row["Site ID (direct)"]
        if row["Site ID (direct)"]
        else (row["Site ID (-0)"] if row["Site ID (-0)"] else row["Site ID (-1)"]),
        axis=1,
    )
    grouped["Match"] = grouped.apply(
        lambda row: "TRUE"
        if row["Match (direct)"] == "TRUE"
        or row["Match -0"] == "TRUE"
        or row["Match -1"] == "TRUE"
        else "FALSE",
        axis=1,
    )

    matched_site_keys: set[str] = set()
    for _, row in grouped.iterrows():
        if row["Match (direct)"] == "TRUE":
            matched_site_keys.add(row["Compare Key (direct)"])
        if row["Match -0"] == "TRUE":
            matched_site_keys.add(row["Compare Key -0"])
        if row["Match -1"] == "TRUE":
            matched_site_keys.add(row["Compare Key -1"])

    unmatched_site_ids: list[str] = []
    seen_unmatched_keys: set[str] = set()
    for value in site_df[site_id_col].tolist():
        original = to_clean_text(value)
        normalized = normalize_id_value(original)
        if not normalized:
            continue
        if normalized in matched_site_keys:
            continue
        if normalized in seen_unmatched_keys:
            continue
        unmatched_site_ids.append(original)
        seen_unmatched_keys.add(normalized)

    unmatched_site_df = pd.DataFrame({"Site ID": unmatched_site_ids})

    output_columns = [
        "Cluster ID",
        "Circle",
        "Site ID",
        "Cluster ID + -0",
        "Cluster ID + -1",
        "Match",
    ]
    grouped = grouped[output_columns]

    matched_df = grouped[grouped["Match"] == "TRUE"].copy()
    unmatched_df = grouped[grouped["Match"] == "FALSE"].copy()

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file) as writer:
        grouped.to_excel(writer, sheet_name="All_ClusterIDs", index=False)
        matched_df.to_excel(writer, sheet_name="Matched", index=False)
        unmatched_df.to_excel(writer, sheet_name="Unmatched", index=False)
        unmatched_site_df.to_excel(writer, sheet_name="Unmatched_SiteIDs", index=False)

    print(f"Comparison complete. Output file: {output_file}")
    print(f"Unique Cluster IDs checked: {len(grouped)}")
    print(f"Matched Cluster IDs: {len(matched_df)}")
    print(f"Unmatched Cluster IDs: {len(unmatched_df)}")
    print(f"Unmatched Site IDs: {len(unmatched_site_df)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
