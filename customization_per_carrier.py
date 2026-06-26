"""Carrier-specific rules for shipment columns, values, and cost display names."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable

import pandas as pd


def normalize_column_name(name: str) -> str:
    return re.sub(r"\s+", " ", str(name).strip().lower())


def _default_shipment_columns() -> list[str]:
    from transform_ratebook import FRONT_COLUMN_ORDER

    return list(FRONT_COLUMN_ORDER)


def _kuehne_nagel_carrier_name(series: pd.Series) -> pd.Series:
    return series.astype(str).str.replace(r"\bKN\b", "Kuehne Nagel", regex=True)


def _kuehne_nagel_carrier_name_from_df(df: pd.DataFrame) -> pd.Series:
    names = _kuehne_nagel_carrier_name(df["Carrier Name"])
    dest = df["Destination Country"].astype(str).str.strip().str.upper()
    return names.mask(dest == "KR", "Kuehne Nagel MA")


# Applied to all carriers (renaming + no green highlight).
COMMON_DISPLAY_NAME_OVERRIDES: list[tuple[str, str]] = [
    ("pre carriage linehaul charge", "Pre-Carriage Linehaul"),
    (
        "pre carriage handling charge",
        "Admin handling Fee (Pre Carriage Handling Charge)",
    ),
    (
        "origin customs clearance fee",
        "Export customs clearance(Origin Customs Clearance Fee)",
    ),
    ("pre carriage thc", "THC origin"),
    ("main carriage non stackable fee", "Non-stackable shipment"),
    ("main carriage rate", "Transport cost"),
    ("pss", "Peak Season Surcharge"),
    ("pps", "Peak Season Surcharge"),
    ("on carriage thc", "THC destination"),
    (
        "on carriage handling charge",
        "Destination handling fee (On Carriage Handling Charge)",
    ),
    ("destination customs clearance fee", "Import customs clearance"),
    (
        "destination documentation turnover fee",
        "Documentation (Destination Documentation Turnover Fee)",
    ),
    ("on carriage linehaul charge", "On-carriage Linehaul"),
]


@dataclass
class CarrierCustomization:
    """Rules applied when the workbook matches one of carrier_keys."""

    carrier_keys: tuple[str, ...]
    shipment_columns: list[str] | None = None
    prepend_transport_mode: bool = False
    exclude_shipment_columns: tuple[str, ...] = ()
    shipment_values: dict[str, str | Callable[[pd.DataFrame], pd.Series]] | None = None
    display_name_overrides: list[tuple[str, str]] | None = None
    extra_display_name_overrides: list[tuple[str, str]] = field(default_factory=list)
    replace_default_display_names: bool = False
    applies_if_by_display_name: dict[str, str] = field(default_factory=dict)
    rate_by_by_display_name: dict[str, str] = field(default_factory=dict)
    skip_rate_columns: set[str] = field(default_factory=set)


CARRIER_CUSTOMIZATIONS: list[CarrierCustomization] = [
    CarrierCustomization(
        carrier_keys=("kn", "kuehne", "kuehne nagel"),
        prepend_transport_mode=True,
        exclude_shipment_columns=("Incoterm",),
        shipment_values={
            "Transport Mode": "AIR",
            "Carrier Name": lambda df: _kuehne_nagel_carrier_name_from_df(df),
        },
        extra_display_name_overrides=[("pss", "Peak Season Surcharge")],
        applies_if_by_display_name={
            "Peak Season Surcharge": "Applies if invoiced by Carrier",
        },
    ),
    CarrierCustomization(
        carrier_keys=("dsv",),
        prepend_transport_mode=True,
        exclude_shipment_columns=("Incoterm",),
        shipment_values={"Transport Mode": "AIR"},
    ),
    CarrierCustomization(
        carrier_keys=("dachser",),
    ),
]


def detect_carrier_key(df: pd.DataFrame) -> str:
    if "Carrier Name" not in df.columns:
        return "default"

    carrier_names = df["Carrier Name"].dropna().astype(str).unique()
    if len(carrier_names) == 0:
        return "default"

    haystack = normalize_column_name(" ".join(carrier_names[:10]))
    for customization in CARRIER_CUSTOMIZATIONS:
        for key in customization.carrier_keys:
            if key != "*" and key in haystack:
                return key

    first = normalize_column_name(str(carrier_names[0]).split()[0])
    return first or "default"


def get_carrier_customization(carrier_key: str) -> CarrierCustomization | None:
    key = normalize_column_name(carrier_key)
    for customization in CARRIER_CUSTOMIZATIONS:
        for carrier_match in customization.carrier_keys:
            if carrier_match == "*" or carrier_match == key:
                return customization
    return None


def _excluded_shipment_column_names(carrier_key: str) -> set[str]:
    customization = get_carrier_customization(carrier_key)
    if not customization or not customization.exclude_shipment_columns:
        return set()
    return {normalize_column_name(name) for name in customization.exclude_shipment_columns}


def get_shipment_columns(carrier_key: str) -> list[str]:
    if normalize_column_name(carrier_key) == "dachser":
        from dachser_customization import dachser_shipment_columns

        columns = dachser_shipment_columns()
    else:
        customization = get_carrier_customization(carrier_key)
        if customization and customization.shipment_columns:
            columns = list(customization.shipment_columns)
        elif customization and customization.prepend_transport_mode:
            columns = ["Transport Mode", *_default_shipment_columns()]
        else:
            columns = _default_shipment_columns()

    excluded = _excluded_shipment_column_names(carrier_key)
    return [col for col in columns if normalize_column_name(col) not in excluded]


def get_display_name_overrides(carrier_key: str) -> list[tuple[str, str]]:
    overrides = list(COMMON_DISPLAY_NAME_OVERRIDES)
    customization = get_carrier_customization(carrier_key)
    if customization and customization.extra_display_name_overrides:
        overrides = list(customization.extra_display_name_overrides) + overrides
    return overrides


def get_standard_display_names(carrier_key: str) -> set[str]:
    """Standard cost labels — never green-highlighted (common + carrier-specific)."""
    names = {display for _, display in get_display_name_overrides(carrier_key)}
    if normalize_column_name(carrier_key) == "dachser":
        from dachser_customization import dachser_destination_display_names

        names |= dachser_destination_display_names()
    return names


# Substrings matched against normalized column names (excluded from final output).
GLOBAL_RATE_COLUMN_SUBSTRING_SKIP = ("transit time",)


def get_skip_rate_columns(carrier_key: str) -> set[str]:
    base = {"currency", "additional request", "category"}
    customization = get_carrier_customization(carrier_key)
    if customization and customization.skip_rate_columns:
        base = base | {normalize_column_name(c) for c in customization.skip_rate_columns}
    return base


def is_excluded_rate_column(column: str, carrier_key: str = "default") -> bool:
    norm = normalize_column_name(column)
    if any(pattern in norm for pattern in GLOBAL_RATE_COLUMN_SUBSTRING_SKIP):
        return True
    return norm in get_skip_rate_columns(carrier_key)


def apply_display_name_override(
    name: str,
    columns: list[str],
    carrier_key: str,
) -> str:
    normalized_name = normalize_column_name(name)
    for token in ("pss", "pps"):
        if (
            normalized_name == token
            or normalized_name.endswith(f" {token}")
            or any(token in normalize_column_name(column) for column in columns)
        ):
            for key, display_name in get_display_name_overrides(carrier_key):
                if key == token:
                    return display_name

    haystack = " ".join(
        [normalized_name] + [normalize_column_name(column) for column in columns]
    )
    for key, display_name in get_display_name_overrides(carrier_key):
        if all(word in haystack for word in key.split()):
            return display_name
    return name


def resolve_applies_if(
    display_name: str,
    columns: list[str],
    default: str,
    carrier_key: str,
) -> str:
    customization = get_carrier_customization(carrier_key)
    if customization and display_name in customization.applies_if_by_display_name:
        return customization.applies_if_by_display_name[display_name]
    return default


def resolve_rate_by(
    display_name: str,
    columns: list[str],
    default: str,
    carrier_key: str,
) -> str:
    customization = get_carrier_customization(carrier_key)
    if customization and display_name in customization.rate_by_by_display_name:
        return customization.rate_by_by_display_name[display_name]
    return default


def apply_shipment_customization(df: pd.DataFrame, carrier_key: str) -> pd.DataFrame:
    """Apply carrier-specific shipment column order and constant/calculated values."""
    customization = get_carrier_customization(carrier_key)
    if customization is None:
        return df

    result = df.copy()
    shipment_cols = get_shipment_columns(carrier_key)
    excluded = _excluded_shipment_column_names(carrier_key)
    drop_cols = [
        col for col in result.columns if normalize_column_name(col) in excluded
    ]
    if drop_cols:
        result = result.drop(columns=drop_cols)

    for column in shipment_cols:
        if column not in result.columns:
            result[column] = pd.NA

    if customization.shipment_values:
        for column, value in customization.shipment_values.items():
            if callable(value):
                result[column] = value(result)
            else:
                result[column] = value

    if normalize_column_name(carrier_key) == "dachser":
        from dachser_customization import apply_dachser_transformations

        result = apply_dachser_transformations(result)
        shipment_cols = get_shipment_columns(carrier_key)

    rate_cols = [col for col in result.columns if col not in shipment_cols]
    available_shipment = [col for col in shipment_cols if col in result.columns]
    return result[available_shipment + rate_cols]
