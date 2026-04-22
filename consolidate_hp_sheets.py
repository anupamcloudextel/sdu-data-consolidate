from __future__ import annotations

from pathlib import Path
import argparse
import sys

import pandas as pd


def normalize_column_name(column_name: object) -> str:
    text = str(column_name).strip().lower()
    return "".join(text.split())


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
            aligned_data[canonical_name] = dataframe[source_column]
        else:
            aligned_data[canonical_name] = pd.Series([pd.NA] * len(dataframe), index=dataframe.index)

    return pd.DataFrame(aligned_data, index=dataframe.index)


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
) -> tuple[pd.DataFrame, list[str]]:
    collected: list[pd.DataFrame] = []
    skipped_files: list[str] = []

    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df["source_file_name"] = file_path.name
            collected.append(df)
            print(f"Included: {file_path.name} ({len(df)} rows)")
        except ValueError:
            skipped_files.append(file_path.name)
            print(f"Skipped (sheet '{sheet_name}' not found): {file_path.name}")
        except Exception as exc:
            skipped_files.append(file_path.name)
            print(f"Skipped (read error): {file_path.name} -> {exc}")

    if collected:
        canonical_order, normalized_to_canonical = build_unified_columns(collected)
        standardized = [
            standardize_dataframe_columns(df, canonical_order, normalized_to_canonical)
            for df in collected
        ]
        consolidated = pd.concat(standardized, ignore_index=True)

        # Keep source file column at the very beginning in output.
        source_column = "source_file_name"
        ordered_columns = [source_column] + [
            column for column in consolidated.columns if column != source_column
        ]
        consolidated = consolidated[ordered_columns]
        return consolidated, skipped_files

    return pd.DataFrame(), skipped_files


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

    consolidated_df, skipped_files = consolidate_sheet_from_files(excel_files, sheet_name)

    if consolidated_df.empty:
        print(
            f"No data consolidated. Either sheet '{sheet_name}' was missing "
            "or files could not be read."
        )
        return 1

    output_file.parent.mkdir(parents=True, exist_ok=True)
    consolidated_df.to_excel(output_file, index=False)

    print("\nConsolidation complete.")
    print(f"Output file: {output_file}")
    print(f"Total input files found: {len(excel_files)}")
    print(f"Files skipped: {len(skipped_files)}")
    print(f"Total consolidated rows: {len(consolidated_df)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
