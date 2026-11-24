"""Pure DataFrame cleaning, filtering, and sub-case aggregation for pier forces.

This module has zero COM/ETABS dependencies and operates strictly on pandas
DataFrames and config dataclasses. Full floating-point precision is preserved
through cleaning so NBCC load combination superposition avoids premature
rounding errors.
"""

from __future__ import annotations

from typing import Mapping, Sequence

import pandas as pd

from src.config import CleaningOptions, NBCC_CATEGORIES, normalize_nbcc_category

REQUIRED_RAW_COLUMNS: tuple[str, ...] = (
    "Story",
    "Pier",
    "LoadCase",
    "Location",
    "P",
    "V2",
    "V3",
    "T",
    "M2",
    "M3",
)
FORCE_COLUMNS: tuple[str, ...] = ("P", "V2", "V3", "T", "M2", "M3")
CLEANED_COLUMNS: tuple[str, ...] = (
    "Story",
    "Pier",
    "LoadCase",
    "Category",
    "Location",
    "P",
    "V2",
    "V3",
    "T",
    "M2",
    "M3",
)

GRAVITY_CATEGORIES: frozenset[str] = frozenset({"Dead Load", "Live Load", "Snow Load"})
LATERAL_CATEGORIES: frozenset[str] = frozenset({"Wind Load", "Seismic Load"})
_LOCATION_ORDER: dict[str, int] = {"Top": 0, "Bottom": 1}
_CATEGORY_ORDER: dict[str, int] = {cat: idx for idx, cat in enumerate(NBCC_CATEGORIES)}


def round_force_columns(df: pd.DataFrame, decimals: int) -> pd.DataFrame:
    """Return a copy of df with force/moment columns rounded to `decimals` places."""
    if decimals < 0:
        raise ValueError(f"decimals must be non-negative, got {decimals}.")
    out = df.copy()
    present = [col for col in FORCE_COLUMNS if col in out.columns]
    if present and not out.empty:
        out[present] = out[present].round(decimals)
        # Normalize -0.0 to 0.0 for clean tabular/Excel output.
        for col in present:
            out[col] = out[col].where(out[col] != 0.0, 0.0)
    return out


def _empty_cleaned_df() -> pd.DataFrame:
    df = pd.DataFrame(columns=list(CLEANED_COLUMNS))
    for col in FORCE_COLUMNS:
        df[col] = df[col].astype(float)
    return df


def _build_order_map(
    preferred_order: Sequence[str] | None,
    observed_values: Sequence[str],
) -> dict[str, int]:
    order: dict[str, int] = {}
    if preferred_order:
        for val in preferred_order:
            if val not in order:
                order[val] = len(order)
    for val in observed_values:
        if val not in order:
            order[val] = len(order)
    return order


def clean_pier_forces(
    raw_df: pd.DataFrame,
    case_mappings: Mapping[str, str],
    options: CleaningOptions,
    *,
    story_order: Sequence[str] | None = None,
    pier_order: Sequence[str] | None = None,
    apply_station_filter: bool = True,
) -> pd.DataFrame:
    """Clean and filter raw ETABS pier force rows prior to NBCC superposition.

    Pipeline:
    1. Validate required columns and coerce force/moment values to float64.
    2. Filter to mapped load cases and assign NBCC categories (D/L/S/W/E).
    3. Filter by Pier, Story, and Station (Location).
    4. Optionally drop exact duplicate rows (preserving distinct multi-step rows).
    5. Detect multi-step rows sharing (Story, Pier, LoadCase, Location) and label
       them `<LoadCase>-1`, `<LoadCase>-2`, ... so steps are never summed together.
    6. Optionally aggregate distinct gravity sub-cases (D, L, S) mapped to the same
       category; lateral cases (W, E) and multi-step cases are always kept separate.
    7. Optionally zero axial P for lateral cases and/or shear & moments for gravity
       cases (reference-notebook cleanup options).
    8. Optionally drop rows where all force/moment magnitudes are below threshold.
    9. Sort deterministically preserving model Story and Pier order.
    """
    options.validate()

    missing_cols = [c for c in REQUIRED_RAW_COLUMNS if c not in raw_df.columns]
    if missing_cols:
        raise ValueError(
            f"Raw pier force DataFrame is missing required column(s): {missing_cols}"
        )

    if raw_df.empty or not case_mappings:
        return _empty_cleaned_df()

    df = raw_df.loc[:, list(REQUIRED_RAW_COLUMNS)].copy()

    # Normalize string identifier columns.
    for col in ("Story", "Pier", "LoadCase", "Location"):
        df[col] = df[col].astype(str).str.strip()

    # Capture natural appearance order before filtering.
    story_rank_map = _build_order_map(story_order, df["Story"].tolist())
    pier_rank_map = _build_order_map(pier_order, df["Pier"].tolist())

    # Coerce force columns strictly so corrupt non-numeric data is not masked.
    for col in FORCE_COLUMNS:
        try:
            df[col] = pd.to_numeric(df[col], errors="raise").astype(float)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Non-numeric value encountered in force column {col!r}."
            ) from exc

    # 1. Keep only mapped load cases and attach canonical NBCC category.
    canonical_mappings = {
        str(k).strip(): normalize_nbcc_category(v) for k, v in case_mappings.items()
    }
    df = df[df["LoadCase"].isin(canonical_mappings.keys())].copy()
    if df.empty:
        return _empty_cleaned_df()

    df["Category"] = df["LoadCase"].map(canonical_mappings)

    # 2. Apply Pier, Story, and optional Station filters.
    if options.selected_piers is not None:
        df = df[df["Pier"].isin(options.selected_piers)].copy()
    if options.selected_stories is not None:
        df = df[df["Story"].isin(options.selected_stories)].copy()
    if apply_station_filter and options.station_filter != "All":
        df = df[df["Location"] == options.station_filter].copy()

    if df.empty:
        return _empty_cleaned_df()

    # 3. Optional exact duplicate row removal (keeps distinct multi-step rows intact).
    if options.remove_duplicates:
        df = df.drop_duplicates(
            subset=["Story", "Pier", "LoadCase", "Location", *FORCE_COLUMNS],
            keep="first",
        ).copy()

    # 4. Preserve multi-step rows (e.g., 3-step seismic eccentricity from PierForce)
    #    as distinct `<LoadCase>-1`, `<LoadCase>-2`, ... labels rather than summing.
    step_key = ["Story", "Pier", "LoadCase", "Location"]
    step_counts = df.groupby(step_key, sort=False)["LoadCase"].transform("size")
    is_multistep = step_counts > 1
    df["_is_multistep"] = is_multistep
    if is_multistep.any():
        step_idx = df.loc[is_multistep].groupby(step_key, sort=False).cumcount() + 1
        df.loc[is_multistep, "LoadCase"] = (
            df.loc[is_multistep, "LoadCase"] + "-" + step_idx.astype(str)
        )

    # 5. Sub-case aggregation for non-stepped gravity categories (D, L, S).
    #    Lateral cases (W, E) and multi-step cases are NEVER summed together.
    if options.aggregate_subcases:
        agg_mask = df["Category"].isin(GRAVITY_CATEGORIES) & (~df["_is_multistep"])
        to_agg = df.loc[agg_mask].copy()
        keep_separate = df.loc[~agg_mask].copy()

        if not to_agg.empty:
            to_agg = (
                to_agg.groupby(
                    ["Story", "Pier", "Category", "Location"],
                    as_index=False,
                    sort=False,
                )[list(FORCE_COLUMNS)]
                .sum()
            )
            to_agg["LoadCase"] = to_agg["Category"]
            to_agg = to_agg[list(CLEANED_COLUMNS)]

        keep_separate = keep_separate[list(CLEANED_COLUMNS)]
        if to_agg.empty:
            df = keep_separate.reset_index(drop=True)
        elif keep_separate.empty:
            df = to_agg.reset_index(drop=True)
        else:
            df = pd.concat([to_agg, keep_separate], ignore_index=True)
    else:
        df = df[list(CLEANED_COLUMNS)].reset_index(drop=True)

    # 6. Optional reference-notebook component zeroing.
    if options.zero_lateral_axial:
        df.loc[df["Category"].isin(LATERAL_CATEGORIES), "P"] = 0.0
    if options.zero_gravity_shear_moment:
        df.loc[
            df["Category"].isin(GRAVITY_CATEGORIES),
            ["V2", "V3", "T", "M2", "M3"],
        ] = 0.0

    # 7. Optional near-zero row filtering.
    if options.drop_near_zero:
        max_abs = df[list(FORCE_COLUMNS)].abs().max(axis=1)
        df = df[max_abs >= options.near_zero_threshold].copy()

    if df.empty:
        return _empty_cleaned_df()

    # 8. Deterministic sort preserving ETABS model story and pier order.
    df["_story_rank"] = df["Story"].map(lambda s: story_rank_map.get(s, 10**9))
    df["_pier_rank"] = df["Pier"].map(lambda p: pier_rank_map.get(p, 10**9))
    df["_loc_rank"] = df["Location"].map(lambda loc: _LOCATION_ORDER.get(loc, 99))
    df["_cat_rank"] = df["Category"].map(lambda c: _CATEGORY_ORDER.get(c, 99))

    df = (
        df.sort_values(
            by=["_story_rank", "_pier_rank", "_loc_rank", "_cat_rank", "LoadCase"],
            kind="mergesort",
        )
        .drop(columns=["_story_rank", "_pier_rank", "_loc_rank", "_cat_rank"])
        .reset_index(drop=True)
    )

    return df[list(CLEANED_COLUMNS)]
