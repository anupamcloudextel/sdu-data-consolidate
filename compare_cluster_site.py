"""
Normalize Cluster ID (consolidated_hp.xlsx) and Site ID (Site.xlsx) for
comparison only, then export trimmed original values to Excel.

Normalization (comparison only):
  - Remove all whitespace.
  - Convert to lowercase.
  - For Cluster ID only: remove "-TPT" or trailing "TPT" (case-insensitive).

Comparison rule:
  - Exact match on normalized values.
  - If no exact match, try Cluster ID + "-0" only when the cluster ID
    does not already end with a numeric suffix (-1, -2, -3, -4, etc.).

Export columns show actual values with leading/trailing whitespace trimmed.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

TPT_DASH_PATTERN = re.compile(r"-tpt", re.IGNORECASE)
TPT_SUFFIX_PATTERN = re.compile(r"tpt$", re.IGNORECASE)
NUMERIC_SUFFIX_PATTERN = re.compile(r"-\d+$")

OUTPUT_COLUMNS = [
    "Cluster ID",
    "Site ID",
    "total match",
    "Unmatched",
]


def trim_id(value: object) -> str | None:
    """Return the actual ID with only leading/trailing whitespace removed."""
    if pd.isna(value):
        return None
    trimmed = str(value).strip()
    return trimmed if trimmed else None


def remove_tpt_suffix(value: str) -> str:
    """Strip -TPT or a trailing TPT from a cluster ID."""
    without_dash_tpt = TPT_DASH_PATTERN.sub("", value)
    return TPT_SUFFIX_PATTERN.sub("", without_dash_tpt)


def normalize_id(value: object, *, remove_tpt: bool = False) -> str | None:
    """Normalize a cluster/site ID for comparison only."""
    if pd.isna(value):
        return None

    normalized = re.sub(r"\s+", "", str(value).strip()).lower()
    if not normalized:
        return None
    if remove_tpt:
        normalized = remove_tpt_suffix(normalized)
        if not normalized:
            return None
    return normalized


def suffix_zero_key(cluster_id: str) -> str | None:
    """Build a -0 comparison key when the cluster ID has no numeric suffix."""
    if NUMERIC_SUFFIX_PATTERN.search(cluster_id):
        return None
    return f"{cluster_id}-0"


def build_display_map(
    normalized_series: pd.Series, display_series: pd.Series
) -> dict[str, str]:
    """Map normalized comparison keys to trimmed original display values."""
    display_map: dict[str, str] = {}
    for normalized, display in zip(normalized_series, display_series):
        if normalized is None or pd.isna(normalized):
            continue
        if normalized not in display_map and display is not None:
            display_map[str(normalized)] = display
    return display_map


def load_hp(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    if "Cluster ID" not in df.columns:
        raise ValueError(f"'Cluster ID' column not found in {path}")
    df["_cluster_display"] = df["Cluster ID"].map(trim_id)
    df["_cluster_normalized"] = df["Cluster ID"].map(
        lambda value: normalize_id(value, remove_tpt=True)
    )
    return df


def load_site(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    if "Site ID" not in df.columns:
        raise ValueError(f"'Site ID' column not found in {path}")
    df["_site_display"] = df["Site ID"].map(trim_id)
    df["_site_normalized"] = df["Site ID"].map(normalize_id)
    return df


def build_comparison(hp_df: pd.DataFrame, site_df: pd.DataFrame) -> pd.DataFrame:
    hp_display = build_display_map(hp_df["_cluster_normalized"], hp_df["_cluster_display"])
    site_display = build_display_map(site_df["_site_normalized"], site_df["_site_display"])

    hp_ids = set(hp_df["_cluster_normalized"].dropna())
    site_ids = set(site_df["_site_normalized"].dropna())
    rows: list[dict[str, str]] = []

    for cluster_norm in sorted(hp_ids):
        cluster_actual = hp_display.get(cluster_norm, cluster_norm)

        if cluster_norm in site_ids:
            site_actual = site_display.get(cluster_norm, cluster_norm)
            rows.append(
                {
                    "Cluster ID": cluster_actual,
                    "Site ID": site_actual,
                    "total match": site_actual,
                    "Unmatched": "",
                }
            )
            continue

        suffix_key = suffix_zero_key(cluster_norm)
        if suffix_key and suffix_key in site_ids:
            site_actual = site_display.get(suffix_key, suffix_key)
            rows.append(
                {
                    "Cluster ID": cluster_actual,
                    "Site ID": site_actual,
                    "total match": site_actual,
                    "Unmatched": "",
                }
            )
            continue

        rows.append(
            {
                "Cluster ID": cluster_actual,
                "Site ID": "",
                "total match": "",
                "Unmatched": cluster_actual,
            }
        )

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Normalize Cluster ID and Site ID, then compare them."
    )
    parser.add_argument(
        "--hp-file",
        type=Path,
        default=Path("consolidated_hp.xlsx"),
        help="Path to consolidated_hp.xlsx",
    )
    parser.add_argument(
        "--site-file",
        type=Path,
        default=Path("Site.xlsx"),
        help="Path to Site.xlsx",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("cluster_site_comparison.xlsx"),
        help="Output Excel report path",
    )
    args = parser.parse_args()

    hp_df = load_hp(args.hp_file)
    site_df = load_site(args.site_file)
    comparison = build_comparison(hp_df, site_df)
    comparison.to_excel(args.output, index=False)

    total_matched = (comparison["total match"] != "").sum()
    unmatched = (comparison["Unmatched"] != "").sum()

    print(f"Report written to: {args.output.resolve()}")
    print(f"  Total matched: {total_matched}")
    print(f"  Unmatched cluster IDs: {unmatched}")


if __name__ == "__main__":
    main()
