from __future__ import annotations

from pathlib import Path
import argparse
import sys

import pandas as pd


def normalize_column_name(column_name: object) -> str:
    text = str(column_name).strip().lower()
    return "".join(text.split())


def get_column_by_normalized_name(dataframe: pd.DataFrame, expected_name: str) -> str | None:
    expected = normalize_column_name(expected_name)
    for column in dataframe.columns:
        if normalize_column_name(column) == expected:
            return str(column)
    return None


def infer_circle_from_file_name(file_name: str) -> str:
    known_circles = (
        "Bangalore",
        "Delhi",
        "Kolkata",
        "Mumbai",
        "Pune",
        "Punjab",
        "UPW",
        "Tamil Nadu",
    )
    lower_name = file_name.lower()
    for circle in known_circles:
        if circle.lower() in lower_name:
            return circle
    return ""


def build_unified_columns(dataframes: list[pd.DataFrame]) -> tuple[list[str], dict[str, str]]:
    normalized_to_canonical: dict[str, str] = {}
    canonical_order: list[str] = []

    for df in dataframes:
        for column in df.columns:
            normalized = normalize_column_name(column)
            if normalized not in normalized_to_canonical:
                normalized_to_canonical[normalized] = str(column).strip()
                canonical_order.append(normalized)

    return canonical_order, normalized_to_canonical


def standardize_dataframe_columns(
    dataframe: pd.DataFrame,
    canonical_order: list[str],
    normalized_to_canonical: dict[str, str],
) -> pd.DataFrame:
    normalized_to_source_column: dict[str, object] = {}
    for column in dataframe.columns:
        normalized_to_source_column[normalize_column_name(column)] = column

    aligned_data: dict[str, pd.Series] = {}
    for normalized in canonical_order:
        canonical_name = normalized_to_canonical[normalized]
        source_column = normalized_to_source_column.get(normalized)
        if source_column is not None:
            source_data = dataframe[source_column]
            if isinstance(source_data, pd.DataFrame):
                # Duplicate column labels can occur after promoting row 1 to headers.
                # In that case, take the first matching column as the source series.
                aligned_data[canonical_name] = source_data.iloc[:, 0]
            else:
                aligned_data[canonical_name] = source_data
        else:
            aligned_data[canonical_name] = pd.Series([pd.NA] * len(dataframe), index=dataframe.index)

    return pd.DataFrame(aligned_data, index=dataframe.index)


# Excel column R (0-based index 17) holds Link Type values in the HP sheet layout.
LINK_TYPE_SOURCE_COLUMN_INDEX = 17
LINK_TYPE_OUTPUT_COLUMN = "Link Type"


def assign_link_type_from_column_r(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Map Excel column R values to a unified Link Type column on each row."""
    if len(dataframe.columns) <= LINK_TYPE_SOURCE_COLUMN_INDEX:
        dataframe[LINK_TYPE_OUTPUT_COLUMN] = pd.NA
        return dataframe

    source_column_name = dataframe.columns[LINK_TYPE_SOURCE_COLUMN_INDEX]
    source_data = dataframe.iloc[:, LINK_TYPE_SOURCE_COLUMN_INDEX]
    if isinstance(source_data, pd.DataFrame):
        source_data = source_data.iloc[:, 0]

    existing_link_type_column = get_column_by_normalized_name(
        dataframe, LINK_TYPE_OUTPUT_COLUMN
    )
    if (
        existing_link_type_column is not None
        and existing_link_type_column != source_column_name
    ):
        dataframe.drop(columns=[existing_link_type_column], inplace=True)

    dataframe[LINK_TYPE_OUTPUT_COLUMN] = source_data
    if source_column_name != LINK_TYPE_OUTPUT_COLUMN:
        dataframe.drop(columns=[source_column_name], inplace=True)

    return dataframe


def promote_first_row_as_header_if_needed(dataframe: pd.DataFrame) -> pd.DataFrame:
    has_unnamed_headers = any(
        str(column).strip().lower().startswith("unnamed:")
        for column in dataframe.columns
    )
    if not has_unnamed_headers or dataframe.empty:
        return dataframe

    first_row = dataframe.iloc[0]
    updated_columns: list[str] = []
    for idx, column in enumerate(dataframe.columns):
        candidate = first_row.iloc[idx]
        candidate_text = str(candidate).strip() if not pd.isna(candidate) else ""
        if candidate_text:
            updated_columns.append(candidate_text)
        else:
            updated_columns.append(str(column).strip())

    adjusted = dataframe.iloc[1:].copy()
    adjusted.columns = updated_columns
    adjusted.reset_index(drop=True, inplace=True)
    return adjusted


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read sheet 'HP' from all Excel files in a folder and "
            "consolidate into one Excel file."
        )
    )
    parser.add_argument(
        "--input-folder",
        default="Master Tracker",
        help="Folder containing input Excel files (default: Master Tracker).",
    )
    parser.add_argument(
        "--sheet-name",
        default="HP",
        help="Sheet name to read from each workbook (default: HP).",
    )
    parser.add_argument(
        "--output-file",
        default="consolidated_hp.xlsx",
        help="Output Excel filename (default: consolidated_hp.xlsx).",
    )
    return parser.parse_args()


def get_excel_files(folder: Path) -> list[Path]:
    allowed_suffixes = {".xlsx", ".xlsm", ".xls"}
    files = [
        file_path
        for file_path in folder.iterdir()
        if file_path.is_file()
        and not file_path.name.startswith("~$")
        and file_path.suffix.lower() in allowed_suffixes
    ]
    return sorted(files, key=lambda p: p.name.lower())


def consolidate_sheet_from_files(
    excel_files: list[Path], sheet_name: str
) -> tuple[pd.DataFrame, pd.DataFrame]:
    collected: list[pd.DataFrame] = []
    report_rows: list[dict[str, object]] = []

    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df = promote_first_row_as_header_if_needed(df)
            df = assign_link_type_from_column_r(df)
            inferred_circle = infer_circle_from_file_name(file_path.name)
            existing_circle_column = get_column_by_normalized_name(df, "Circle")
            if existing_circle_column is None:
                df["Circle"] = inferred_circle
            else:
                df["Circle"] = df[existing_circle_column].map(
                    lambda value: str(value).strip() if not pd.isna(value) else ""
                )
                if inferred_circle:
                    df.loc[df["Circle"] == "", "Circle"] = inferred_circle
                if existing_circle_column != "Circle":
                    df.drop(columns=[existing_circle_column], inplace=True)
            df["source_file_name"] = file_path.name
            collected.append(df)
            report_rows.append(
                {
                    "file_name": file_path.name,
                    "status": "SUCCESS",
                    "rows_read": len(df),
                    "rows_failed": 0,
                    "error_reason": "",
                }
            )
            print(f"Included: {file_path.name} ({len(df)} rows)")
        except ValueError:
            report_rows.append(
                {
                    "file_name": file_path.name,
                    "status": "FAILED",
                    "rows_read": 0,
                    "rows_failed": 0,
                    "error_reason": f"Sheet '{sheet_name}' not found",
                }
            )
            print(f"Skipped (sheet '{sheet_name}' not found): {file_path.name}")
        except Exception as exc:
            report_rows.append(
                {
                    "file_name": file_path.name,
                    "status": "FAILED",
                    "rows_read": 0,
                    "rows_failed": 0,
                    "error_reason": str(exc),
                }
            )
            print(f"Skipped (read error): {file_path.name} -> {exc}")

    report_df = pd.DataFrame(report_rows)

    if collected:
        canonical_order, normalized_to_canonical = build_unified_columns(collected)
        standardized = [
            standardize_dataframe_columns(df, canonical_order, normalized_to_canonical)
            for df in collected
        ]
        consolidated = pd.concat(standardized, ignore_index=True)

        # Keep source file column at the very beginning in output, then Circle,
        # and place Link Type immediately after Cluster ID for row-wise readability.
        source_column = "source_file_name"
        circle_column = "Circle"
        link_type_column = LINK_TYPE_OUTPUT_COLUMN
        cluster_id_column = get_column_by_normalized_name(consolidated, "Cluster ID")
        remaining_columns = [
            column
            for column in consolidated.columns
            if column not in (source_column, circle_column, link_type_column)
        ]
        if cluster_id_column is not None and cluster_id_column in remaining_columns:
            insert_at = remaining_columns.index(cluster_id_column) + 1
            remaining_columns.insert(insert_at, link_type_column)
        else:
            remaining_columns.append(link_type_column)
        ordered_columns = [source_column, circle_column] + remaining_columns
        consolidated = consolidated[ordered_columns]
        return consolidated, report_df

    return pd.DataFrame(), report_df


def main() -> int:
    args = parse_args()

    input_folder = Path(args.input_folder).resolve()
    output_file = Path(args.output_file).resolve()
    sheet_name = args.sheet_name

    if not input_folder.exists() or not input_folder.is_dir():
        print(f"Input folder does not exist or is not a directory: {input_folder}")
        return 1

    excel_files = get_excel_files(input_folder)
    if not excel_files:
        print(f"No Excel files found in folder: {input_folder}")
        return 1

    consolidated_df, report_df = consolidate_sheet_from_files(excel_files, sheet_name)

    if consolidated_df.empty:
        print(
            f"No data consolidated. Either sheet '{sheet_name}' was missing "
            "or files could not be read."
        )
        return 1

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file) as writer:
        consolidated_df.to_excel(writer, sheet_name="Consolidated", index=False)
        report_df.to_excel(writer, sheet_name="FileWiseReport", index=False)

    print("\nConsolidation complete.")
    print(f"Output file: {output_file}")
    print(f"Total input files found: {len(excel_files)}")
    failed_files = int((report_df["status"] == "FAILED").sum()) if not report_df.empty else 0
    print(f"Files skipped: {failed_files}")
    print(f"Total consolidated rows: {len(consolidated_df)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
