"""Load a chosen Excel tab from input/ into a pandas DataFrame."""

from pathlib import Path

import pandas as pd

from export_rates_layout import export_rates_layout_xlsx
from transform_ratebook import save_processed_xlsx, transform_ratebook_df

INPUT_DIR = Path(__file__).resolve().parent / "input"


def list_xlsx_files() -> list[Path]:
    files = sorted(
        p for p in INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in {INPUT_DIR}")
    return files


def choose_file(files: list[Path]) -> Path:
    print("Workbooks in input/:")
    for i, path in enumerate(files, start=1):
        print(f"  {i}. {path.name}")

    while True:
        choice = input("Select workbook to process (number): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            selected = files[int(choice) - 1]
            print(f"Using workbook: {selected.name}")
            return selected
        print(f"Enter a number between 1 and {len(files)}.")


def choose_sheet(sheet_names: list[str]) -> str:
    print("\nAvailable tabs:")
    for i, name in enumerate(sheet_names, start=1):
        print(f"  {i}. {name}")

    while True:
        choice = input("Select tab number or name: ").strip()
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(sheet_names):
                return sheet_names[idx - 1]
        elif choice in sheet_names:
            return choice
        print("Invalid choice. Enter a tab number or exact tab name.")


def xlsx_tab_to_df(xlsx_path: Path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(xlsx_path, sheet_name=sheet_name)


def find_lane_id_column(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        if str(col).replace(" ", "").lower() == "laneid":
            return col
    return None


def ask_yes_no(prompt: str) -> bool:
    while True:
        answer = input(f"{prompt} (y/n): ").strip().lower()
        if answer in {"y", "yes"}:
            return True
        if answer in {"n", "no"}:
            return False
        print("Please enter y or n.")


def remove_non_emea_rows(df: pd.DataFrame) -> pd.DataFrame:
    lane_col = find_lane_id_column(df)
    if lane_col is None:
        print("Skipping Non-EMEA cleanup: no Lane ID column found.")
        return df

    before = len(df)
    mask = df[lane_col].astype(str).str.contains("EMEA", case=False, na=False)
    cleaned = df.loc[mask].copy()
    removed = before - len(cleaned)
    print(f"Removed {removed} non-EMEA rows ({before} -> {len(cleaned)}).")
    return cleaned


def main() -> pd.DataFrame:
    print("Select input file to process.\n")
    xlsx_path = choose_file(list_xlsx_files())
    sheet_names = pd.ExcelFile(xlsx_path).sheet_names
    sheet_name = choose_sheet(sheet_names)

    print(f"\nLoading '{sheet_name}' from {xlsx_path.name}...")
    df = xlsx_tab_to_df(xlsx_path, sheet_name)

    if ask_yes_no("Remove Non-EMEA rows (keep only Lane ID containing EMEA)?"):
        df = remove_non_emea_rows(df)

    print("\nTransforming columns and preparing export...")
    df = transform_ratebook_df(df, xlsx_path)
    output_path = save_processed_xlsx(df, xlsx_path, sheet_name)
    print(f"Saved processed workbook: {output_path}")

    layout_path = export_rates_layout_xlsx(df, xlsx_path, sheet_name)
    print(f"Saved rates layout workbook: {layout_path}")

    print(f"DataFrame shape: {df.shape[0]} rows x {df.shape[1]} columns")
    print("\nPreview:")
    print(df.head())

    return df


if __name__ == "__main__":
    df = main()
