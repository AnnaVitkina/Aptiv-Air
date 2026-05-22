"""Column selection, renaming, and ratebook export for the cleaned bid DataFrame."""

import re
from pathlib import Path

import pandas as pd

from customization_per_carrier import (
    apply_shipment_customization,
    detect_carrier_key,
)

PROCESSING_DIR = Path(__file__).resolve().parent / "processing"

FRONT_COLUMN_ORDER = [
    "Paying Region",
    "Lane Id",
    "Origin Region",
    "Origin Country",
    "Origin City",
    "Origin Zip Code",
    "Origin State",
    "Origin Airport",
    "Destination Region",
    "Destination Country",
    "Destination City",
    "Destination Zip Code",
    "Destination State",
    "Destination Airport",
    "Service Level",
    "Carrier Name",
    "Incoterm",
    "Valid from",
    "Valid to",
]


def normalize_column_name(name: str) -> str:
    return re.sub(r"\s+", " ", str(name).strip().lower())


def find_column(df: pd.DataFrame, *candidates: str) -> str:
    lookup = {normalize_column_name(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_column_name(candidate)
        if key in lookup:
            return lookup[key]
    raise KeyError(f"Missing column. Expected one of: {', '.join(candidates)}")


def paying_region_from_lane_id(series: pd.Series) -> pd.Series:
    return series.astype(str).str.extract(r"^([A-Za-z]+)", expand=False)


def format_ratebook_date(value) -> str:
    return pd.to_datetime(value, dayfirst=True).strftime("%d.%m.%Y")


def read_revision_dates(xlsx_path: Path) -> tuple[str, str]:
    revision_raw = pd.read_excel(xlsx_path, sheet_name="Revision", header=None)
    valid_to = None
    for _, row in revision_raw.iterrows():
        for idx, cell in enumerate(row):
            if str(cell).strip().lower() == "until:":
                valid_to = row.iloc[idx + 1]
                break
        if valid_to is not None:
            break
    if valid_to is None or pd.isna(valid_to):
        raise ValueError("Could not find 'until:' date on Revision tab.")

    revision_table = pd.read_excel(xlsx_path, sheet_name="Revision", header=4)
    effective_col = find_column(revision_table, "Effective date")
    effective_dates = revision_table[effective_col].dropna()
    if effective_dates.empty:
        raise ValueError("No Effective date values found on Revision tab.")
    valid_from = effective_dates.iloc[-1]

    return format_ratebook_date(valid_from), format_ratebook_date(valid_to)


def first_pre_carriage_column(columns: pd.Index) -> str | None:
    for col in columns:
        if "pre carriage" in normalize_column_name(col) or "pre-carriage" in normalize_column_name(
            col
        ):
            return col
    return None


def rate_columns_from(df: pd.DataFrame) -> list[str]:
    start_col = first_pre_carriage_column(df.columns)
    if start_col is None:
        raise KeyError("No Pre Carriage columns found in DataFrame.")
    start_idx = df.columns.get_loc(start_col)
    return list(df.columns[start_idx:])


def build_carrier_name(df: pd.DataFrame) -> pd.Series:
    carrier_col = find_column(df, "Carrier Code")
    destination_code_col = find_column(
        df, "Destination Country Code", "destination country code"
    )
    carrier = df[carrier_col].astype(str).str.strip()
    destination = df[destination_code_col].astype(str).str.strip()
    return carrier + " " + destination


def transform_ratebook_df(df: pd.DataFrame, xlsx_path: Path) -> pd.DataFrame:
    lane_col = find_column(df, "Lane ID", "Lane Id")
    valid_from, valid_to = read_revision_dates(xlsx_path)

    front = pd.DataFrame(index=df.index)
    front["Paying Region"] = paying_region_from_lane_id(df[lane_col])
    front["Lane Id"] = df[lane_col]
    front["Origin Region"] = df[find_column(df, "Origin Region")]
    front["Origin Country"] = df[find_column(df, "Origin Country code", "Origin country code")]
    front["Origin City"] = df[find_column(df, "Origin City")]
    front["Origin Zip Code"] = df[
        find_column(df, "Origin Postal Code", "Origin postal code")
    ]
    front["Origin State"] = df[find_column(df, "Origin State")]
    front["Origin Airport"] = df[
        find_column(df, "Proposed Origin Airport", "Proposed origin airport")
    ]
    front["Destination Region"] = df[find_column(df, "Destination Region")]
    front["Destination Country"] = df[
        find_column(df, "Destination Country Code", "destination country code")
    ]
    front["Destination City"] = df[find_column(df, "Destination City", "Destination city")]
    front["Destination Zip Code"] = df[
        find_column(df, "Destination Postal Code", "Destination postal code")
    ]
    front["Destination State"] = df[find_column(df, "Destination State", "destination state")]
    front["Destination Airport"] = df[
        find_column(df, "Proposed Destination Airport", "Proposed Destination Airport")
    ]
    front["Service Level"] = df[find_column(df, "Service")]
    front["Carrier Name"] = build_carrier_name(df)
    front["Incoterm"] = "exc. DAP"
    front["Valid from"] = valid_from
    front["Valid to"] = valid_to

    currency = df[find_column(df, "Currency")]
    rate_cols = rate_columns_from(df)
    result = pd.concat(
        [front[FRONT_COLUMN_ORDER], currency.rename("Currency"), df[rate_cols]],
        axis=1,
    )
    carrier_key = detect_carrier_key(result)
    return apply_shipment_customization(result, carrier_key)


def save_processed_xlsx(df: pd.DataFrame, source_path: Path, sheet_name: str) -> Path:
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    safe_sheet = re.sub(r'[<>:"/\\|?*]', "_", sheet_name)
    output_path = PROCESSING_DIR / f"{source_path.stem}_{safe_sheet}_processed.xlsx"
    df.to_excel(output_path, index=False, sheet_name=safe_sheet[:31])
    return output_path
