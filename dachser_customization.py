"""Dachser-specific row expansion, destination charge columns, and export rules."""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from customization_per_carrier import normalize_column_name
from export_rates_layout import (
    BLOCK_DEST,
    BLOCK_ORDER,
    RateCost,
    RateSubColumn,
    build_sub_columns,
    classify_block,
    group_rate_columns,
    infer_applies_if,
    infer_display_name,
    infer_rate_by,
    rate_columns_from_df,
)
from customization_per_carrier import (
    apply_display_name_override,
    get_shipment_columns,
    resolve_applies_if,
    resolve_rate_by,
)

# Destination country code -> Carrier Name (exact mappings only).
DACHSER_CARRIER_NAME_BY_DESTINATION: dict[str, str] = {
    "DE": "Dachser DE",
    "MK": "Dachser DE",
    "AT": "Dachser AT (PN1)",
    "HU": "Dachser HU",
    "ES": "Dachser ES",
    "MA": "Dachser MA",
    "PL": "Dachser PL",
    "TR": "Dachser TR",
    "PT": "Dachser PT",
}

DACHSER_VAT_CARRIER_COUNTRY: dict[str, str] = {
    "Dachser DE": "DE",
    "Dachser AT (PN1)": "AT",
    "Dachser ES": "ES",
    "Dachser MA": "MA",
    "Dachser TR": "TR",
    "Dachser PT": "PT",
}

DACHSER_PL_HU_CARRIERS = ("Dachser HU", "Dachser PL")
PL_HU_APPLIES_IF = (
    "Applies if: Carrier Name equals 'Dachser HU', 'Dachser PL'"
)
VAT_COUNTRY_ORDER = sorted(set(DACHSER_VAT_CARRIER_COUNTRY.values()))

# Source column token -> standard destination charge display name.
DACHSER_DESTINATION_CHARGE_PATTERNS: list[tuple[str, str]] = [
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

DACHSER_DEST_PREFIX = "__dachser_dest__"


@dataclass(frozen=True)
class DachserDestVariant:
    display_base: str
    display_name: str
    applies_if: str
    suffix: str


def dachser_shipment_columns() -> list[str]:
    from transform_ratebook import FRONT_COLUMN_ORDER

    columns = list(FRONT_COLUMN_ORDER)
    origin_idx = columns.index("Origin Country")
    columns.insert(origin_idx + 1, "VAT")
    return columns


def dachser_carrier_name(df: pd.DataFrame) -> pd.Series:
    dest = df["Destination Country"].astype(str).str.strip().str.upper()
    return dest.map(DACHSER_CARRIER_NAME_BY_DESTINATION)


def _destination_source_columns(df: pd.DataFrame) -> list[str]:
    shipment_names = {
        normalize_column_name(name) for name in get_shipment_columns("dachser")
    }
    return [
        col
        for col in df.columns
        if not str(col).startswith(DACHSER_DEST_PREFIX)
        and normalize_column_name(col) not in shipment_names
        and classify_block(col) == BLOCK_DEST
    ]


def find_destination_source_group(
    df: pd.DataFrame, pattern_key: str, display_base: str
) -> list[str] | None:
    """Find the full original column group for a destination charge (min + p/unit, etc.)."""
    words = pattern_key.split()
    for group in group_rate_columns(_destination_source_columns(df)):
        inferred = apply_display_name_override(
            infer_display_name(group), group, "dachser"
        )
        if inferred == display_base:
            return group
        haystack = " ".join(normalize_column_name(col) for col in group)
        if all(word in haystack for word in words):
            return group
    return None


def _dest_column_name(display_base: str, suffix: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9]+", "_", display_base).strip("_").lower()
    safe_suffix = re.sub(r"[^A-Za-z0-9]+", "_", suffix).strip("_").lower()
    return f"{DACHSER_DEST_PREFIX}{safe}__{safe_suffix}"


def _expanded_column_name(display_base: str, variant_suffix: str, sub_idx: int) -> str:
    return _dest_column_name(display_base, f"{variant_suffix}__{sub_idx}")


def build_dachser_destination_variants() -> list[DachserDestVariant]:
    """Build destination variants: PL/HU block, then each country (non-VAT then VAT)."""
    variants: list[DachserDestVariant] = []
    country_to_carrier = {cc: name for name, cc in DACHSER_VAT_CARRIER_COUNTRY.items()}

    for _pattern, display_base in DACHSER_DESTINATION_CHARGE_PATTERNS:
        variants.append(
            DachserDestVariant(
                display_base=display_base,
                display_name=f"{display_base} (PL, HU)",
                applies_if=PL_HU_APPLIES_IF,
                suffix="pl_hu",
            )
        )

    for country in VAT_COUNTRY_ORDER:
        carrier = country_to_carrier[country]
        for _pattern, display_base in DACHSER_DESTINATION_CHARGE_PATTERNS:
            variants.append(
                DachserDestVariant(
                    display_base=display_base,
                    display_name=f"{display_base} ({country}, non-VAT)",
                    applies_if=(
                        f"Applies if: Carrier Name equals '{carrier}' "
                        f"and LOADING_LOCATION does not equal VAT_{country}"
                    ),
                    suffix=f"{country.lower()}_non_vat",
                )
            )
        for _pattern, display_base in DACHSER_DESTINATION_CHARGE_PATTERNS:
            variants.append(
                DachserDestVariant(
                    display_base=display_base,
                    display_name=f"{display_base} ({country}, VAT)",
                    applies_if=(
                        f"Applies if: Carrier Name equals '{carrier}' "
                        f"and LOADING_LOCATION equals VAT_{country}"
                    ),
                    suffix=f"{country.lower()}_vat",
                )
            )

    return variants


def _variants_by_display_base() -> dict[str, list[DachserDestVariant]]:
    grouped: dict[str, list[DachserDestVariant]] = {}
    for variant in build_dachser_destination_variants():
        grouped.setdefault(variant.display_base, []).append(variant)
    return grouped


def _variant_fill_mask(df: pd.DataFrame, variant: DachserDestVariant) -> pd.Series:
    if "(PL, HU)" in variant.display_name:
        return df["Carrier Name"].isin(DACHSER_PL_HU_CARRIERS) & df["VAT"].isna()

    carrier = _carrier_for_vat_variant(variant)
    if carrier is None:
        return pd.Series(False, index=df.index)

    if ", non-VAT)" in variant.display_name:
        return (df["Carrier Name"] == carrier) & df["VAT"].isna()
    if ", VAT)" in variant.display_name:
        country = DACHSER_VAT_CARRIER_COUNTRY[carrier]
        return (df["Carrier Name"] == carrier) & (df["VAT"] == f"VAT_{country}")
    return pd.Series(False, index=df.index)


def _carrier_for_vat_variant(variant: DachserDestVariant) -> str | None:
    for carrier, country in DACHSER_VAT_CARRIER_COUNTRY.items():
        if f"({country}, non-VAT)" in variant.display_name:
            return carrier
        if f"({country}, VAT)" in variant.display_name:
            return carrier
    return None


def _display_base_from_variant(variant: DachserDestVariant) -> str:
    return variant.display_base


def _dachser_destination_sub_columns(
    variant: DachserDestVariant,
    display_base: str,
    dest_structures: dict[str, list[RateSubColumn]],
) -> list[RateSubColumn]:
    template = dest_structures.get(display_base)
    if not template:
        return [
            RateSubColumn(header="Currency", is_currency=True),
            RateSubColumn(
                header="Flat",
                source_column=_expanded_column_name(display_base, variant.suffix, 0),
            ),
        ]

    sub_columns = [RateSubColumn(header="Currency", is_currency=True)]
    sub_idx = 0
    for item in template:
        if item.is_currency:
            continue
        sub_columns.append(
            RateSubColumn(
                header=item.header,
                source_column=_expanded_column_name(display_base, variant.suffix, sub_idx),
                min_max_label=item.min_max_label,
            )
        )
        sub_idx += 1
    return sub_columns


def _variant_has_columns(
    variant: DachserDestVariant,
    display_base: str,
    dest_structures: dict[str, list[RateSubColumn]],
    df: pd.DataFrame,
) -> bool:
    template = dest_structures.get(display_base, [])
    rate_count = sum(1 for item in template if not item.is_currency)
    if rate_count == 0:
        rate_count = 1
    for sub_idx in range(rate_count):
        if _expanded_column_name(display_base, variant.suffix, sub_idx) in df.columns:
            return True
    return False


def expand_dachser_destination_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    variants_by_base = _variants_by_display_base()
    dest_structures: dict[str, list[RateSubColumn]] = {}
    dest_rate_by: dict[str, str] = {}

    for _pattern, display_base in DACHSER_DESTINATION_CHARGE_PATTERNS:
        source_group = find_destination_source_group(result, _pattern, display_base)
        if source_group is None:
            continue

        sub_columns_template = build_sub_columns(source_group)
        dest_structures[display_base] = sub_columns_template
        dest_rate_by[display_base] = infer_rate_by(source_group)

        source_rate_cols = [
            sub.source_column
            for sub in sub_columns_template
            if not sub.is_currency and sub.source_column
        ]

        for variant in variants_by_base.get(display_base, []):
            mask = _variant_fill_mask(result, variant)
            for sub_idx, source_col in enumerate(source_rate_cols):
                exp_col = _expanded_column_name(display_base, variant.suffix, sub_idx)
                if exp_col not in result.columns:
                    result[exp_col] = pd.NA
                if mask.any():
                    result.loc[mask, exp_col] = result.loc[mask, source_col]

        result = result.drop(columns=source_group)

    result.attrs["dachser_dest_structures"] = dest_structures
    result.attrs["dachser_dest_rate_by"] = dest_rate_by
    return result


def expand_dachser_rows(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["VAT"] = pd.NA

    duplicate_mask = result["Carrier Name"].isin(DACHSER_VAT_CARRIER_COUNTRY)
    duplicates = result.loc[duplicate_mask].copy()
    duplicates["VAT"] = duplicates["Carrier Name"].map(
        {name: f"VAT_{cc}" for name, cc in DACHSER_VAT_CARRIER_COUNTRY.items()}
    )

    return pd.concat([result, duplicates], ignore_index=True)


def clear_non_destination_rates_on_vat_rows(df: pd.DataFrame) -> pd.DataFrame:
    """VAT duplicate rows carry destination VAT costs only — not origin/main freight."""
    result = df.copy()
    vat_mask = result["VAT"].notna()
    if not vat_mask.any():
        return result

    shipment_cols = set(dachser_shipment_columns()) | {"Currency"}
    for col in result.columns:
        if col in shipment_cols or _is_dachser_destination_column(col):
            continue
        if classify_block(col) != BLOCK_DEST:
            result.loc[vat_mask, col] = pd.NA
    return result


def apply_dachser_transformations(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["Carrier Name"] = dachser_carrier_name(result)
    result = expand_dachser_rows(result)
    result = expand_dachser_destination_columns(result)
    result = clear_non_destination_rates_on_vat_rows(result)
    return result


def _is_dachser_destination_column(column: str) -> bool:
    return str(column).startswith(DACHSER_DEST_PREFIX)


def build_dachser_rate_cost_groups(df: pd.DataFrame, carrier_key: str) -> list[RateCost]:
    columns = rate_columns_from_df(df, carrier_key)
    # Dachser PL/HU/VAT rules apply only to destination charges; keep origin/main standard.
    origin_main_columns = [
        col
        for col in columns
        if not _is_dachser_destination_column(col) and classify_block(col) != BLOCK_DEST
    ]

    by_block: dict[str, list[RateCost]] = {block: [] for block in BLOCK_ORDER}

    for group_columns in group_rate_columns(origin_main_columns):
        block = classify_block(group_columns[0])
        inferred_name = infer_display_name(group_columns)
        display_name = apply_display_name_override(
            inferred_name, group_columns, carrier_key
        )
        by_block[block].append(
            RateCost(
                block=block,
                name=display_name,
                applies_if=resolve_applies_if(
                    display_name,
                    group_columns,
                    infer_applies_if(group_columns),
                    carrier_key,
                ),
                rate_by=resolve_rate_by(
                    display_name,
                    group_columns,
                    infer_rate_by(group_columns),
                    carrier_key,
                ),
                sub_columns=build_sub_columns(group_columns),
            )
        )

    dest_structures: dict[str, list[RateSubColumn]] = df.attrs.get(
        "dachser_dest_structures", {}
    )
    dest_rate_by: dict[str, str] = df.attrs.get("dachser_dest_rate_by", {})

    destination_costs: list[RateCost] = []
    for variant in build_dachser_destination_variants():
        display_base = _display_base_from_variant(variant)
        if display_base not in dest_structures:
            continue
        if not _variant_has_columns(variant, display_base, dest_structures, df):
            continue
        destination_costs.append(
            RateCost(
                block=BLOCK_DEST,
                name=variant.display_name,
                applies_if=variant.applies_if,
                rate_by=dest_rate_by.get(display_base, ""),
                sub_columns=_dachser_destination_sub_columns(
                    variant, display_base, dest_structures
                ),
            )
        )

    ordered_origin_main: list[RateCost] = []
    for block in BLOCK_ORDER:
        if block != BLOCK_DEST:
            ordered_origin_main.extend(by_block[block])

    return ordered_origin_main + destination_costs


def dachser_destination_display_names() -> set[str]:
    return {variant.display_name for variant in build_dachser_destination_variants()}
