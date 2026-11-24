"""Declarative NBCC 2020 Table 4.1.3.2.-A ULS Load Combinations & Governing Envelope.

Reference:
    National Building Code of Canada (NBCC 2020), Division B, Part 4,
    Table 4.1.3.2.-A — Load Combinations Without Crane Loads for Ultimate Limit States:

        Case 1: 1.4D
        Case 2: (1.25D or 0.9D) + 1.5L  [+ 1.0S  or  +/- 0.4W]
        Case 3: (1.25D or 0.9D) + 1.5S  [+ 1.0L  or  +/- 0.4W]
        Case 4: (1.25D or 0.9D) +/- 1.4W [+ 0.5L (1.0L storage)  or  + 0.5S]
        Case 5: 1.0D +/- 1.0E            [+ 0.5L (1.0L storage) + 0.25S]

    Per NBCC 2020 Clause 4.1.3.2.(2), companion loads are omitted when they
    reduce the load effect, so both principal-only and companion variants are
    evaluated whenever companion loads are provided.
    Per NBCC 2020 Clause 4.1.3.2.(7) / Table 4.1.3.2.-A Note (6), the companion
    live-load factor 0.5L in Cases 4 and 5 increases to 1.0L for storage,
    equipment, and service areas (Case 3 already uses 1.0L for all occupancies).

This module has zero COM/ETABS dependencies and accepts/returns pure pandas
DataFrames.
"""

from __future__ import annotations

from dataclasses import dataclass
from itertools import product
from typing import Sequence

import pandas as pd

from src.cleaner import FORCE_COLUMNS, round_force_columns
from src.config import (
    CleaningOptions,
    ComboDefinition,
    NBCC_CATEGORIES,
    NBCC_CATEGORY_SYMBOLS,
    NBCCOptions,
    normalize_nbcc_category,
)

COMBO_OUTPUT_COLUMNS: tuple[str, ...] = (
    "Story",
    "Pier",
    "Location",
    "ComboName",
    "P",
    "T",
    "V2",
    "M3",
    "Cm3",
    "V3",
    "M2",
    "Cm2",
)

ENVELOPE_EXTREME_COMPONENTS: tuple[str, ...] = ("P", "V2", "V3", "M2", "M3")
ENVELOPE_COLUMNS: tuple[str, ...] = (
    "Pier",
    "Story",
    "Combo",
    "Extreme",
    "Location",
    "P",
    "T",
    "V2",
    "M3",
    "Cm3",
    "V3",
    "M2",
    "Cm2",
)
_ENVELOPE_FORCE_COLUMNS: tuple[str, ...] = ("P", "T", "V2", "M3", "V3", "M2")
CM_COLUMNS: tuple[str, ...] = ("Cm3", "Cm2")
_CM_ZERO_EPS: float = 1e-12


def compute_cm_factor(m_top: float | None, m_bot: float | None) -> float:
    """Compute CSA A23.3 Clause 10.15.3.1 equivalent uniform moment factor Cm.

    Formula:
        Cm = max(0.4, 0.6 + 0.4 * (M1 / M2))

    where |M1| <= |M2| are the factored end moments at Top and Bottom of the pier.
    Under ETABS pier local-axis sign convention:
      - Same sign (m_top * m_bot > 0): single curvature -> M1 / M2 > 0
      - Opposite sign (m_top * m_bot < 0): double curvature -> M1 / M2 < 0
      - Both moments effectively zero or opposite end unavailable: conservative 1.0
    """
    if m_top is None or m_bot is None:
        return 1.0
    if pd.isna(m_top) or pd.isna(m_bot):
        return 1.0

    abs_top = abs(m_top)
    abs_bot = abs(m_bot)
    if max(abs_top, abs_bot) <= _CM_ZERO_EPS:
        return 1.0

    if abs_top <= abs_bot:
        ratio = m_top / m_bot
    else:
        ratio = m_bot / m_top

    return max(0.4, 0.6 + 0.4 * ratio)


def _attach_cm_columns(combos_df: pd.DataFrame) -> pd.DataFrame:
    """Attach strong-axis `Cm3` and weak-axis `Cm2` using paired Top/Bottom moments."""
    out = combos_df.copy()
    if out.empty:
        out["Cm3"] = pd.Series(dtype=float)
        out["Cm2"] = pd.Series(dtype=float)
        return out

    # Map (Story, Pier, ComboName) -> {"Top": (M3, M2), "Bottom": (M3, M2)}
    paired: dict[tuple[str, str, str], dict[str, tuple[float, float]]] = {}
    for row in out.itertuples(index=False):
        key = (str(row.Story), str(row.Pier), str(row.ComboName))
        loc = str(row.Location)
        paired.setdefault(key, {})[loc] = (float(row.M3), float(row.M2))

    cm3_vals: list[float] = []
    cm2_vals: list[float] = []
    for row in out.itertuples(index=False):
        key = (str(row.Story), str(row.Pier), str(row.ComboName))
        stations = paired.get(key, {})
        top = stations.get("Top")
        bot = stations.get("Bottom")
        m3_top = top[0] if top is not None else None
        m3_bot = bot[0] if bot is not None else None
        m2_top = top[1] if top is not None else None
        m2_bot = bot[1] if bot is not None else None
        cm3_vals.append(compute_cm_factor(m3_top, m3_bot))
        cm2_vals.append(compute_cm_factor(m2_top, m2_bot))

    out["Cm3"] = cm3_vals
    out["Cm2"] = cm2_vals
    return out


def _round_cm_columns(df: pd.DataFrame, decimals: int) -> pd.DataFrame:
    """Round Cm columns to at least 2 decimal places so factors never collapse to 0/1."""
    cm_dec = max(2, decimals)
    out = df.copy()
    for col in CM_COLUMNS:
        if col in out.columns and not out.empty:
            out[col] = out[col].round(cm_dec)
    return out


@dataclass(frozen=True)
class ComboSpec:
    """Declarative NBCC 2020 ULS load combination formula specification."""

    case_id: str
    terms: tuple[tuple[str, float], ...]  # Ordered (NBCC category, load factor)

    @property
    def categories(self) -> frozenset[str]:
        return frozenset(cat for cat, _ in self.terms)


def build_nbcc_2020_specs(
    present_categories: set[str] | frozenset[str],
    nbcc_options: NBCCOptions,
) -> list[ComboSpec]:
    """Build the ordered, declarative list of NBCC 2020 Table 4.1.3.2.-A specs
    applicable to `present_categories`.

    Every spec returned references ONLY categories present in `present_categories`,
    guaranteeing no phantom or misleading combinations are emitted.
    """
    cats = frozenset(normalize_nbcc_category(c) for c in present_categories)
    has_d = "Dead Load" in cats
    has_l = "Live Load" in cats
    has_s = "Snow Load" in cats
    has_w = "Wind Load" in cats
    has_e = "Seismic Load" in cats

    # Companion live load factor for Cases 4 and 5 (NBCC 2020 Clause 4.1.3.2.(7)).
    c_l = 1.0 if nbcc_options.storage_assembly_occupancy else 0.5

    dead_variants: tuple[tuple[tuple[str, float], ...], ...] = (
        ((("Dead Load", 1.25),), (("Dead Load", 0.9),)) if has_d else ((),)
    )
    dead_seismic: tuple[tuple[str, float], ...] = (
        (("Dead Load", 1.0),) if has_d else ()
    )

    specs: list[ComboSpec] = []

    # ------------------------------------------------------------------
    # Case 1: 1.4D
    # ------------------------------------------------------------------
    if has_d:
        specs.append(ComboSpec("Case 1", (("Dead Load", 1.4),)))

    # ------------------------------------------------------------------
    # Case 2: (1.25D or 0.9D) + 1.5L [+ 1.0S or +/- 0.4W]
    # ------------------------------------------------------------------
    if has_l:
        for d_terms in dead_variants:
            base = d_terms + (("Live Load", 1.5),)
            specs.append(ComboSpec("Case 2", base))
            if has_s:
                specs.append(ComboSpec("Case 2 (+S)", base + (("Snow Load", 1.0),)))
            if has_w:
                specs.append(ComboSpec("Case 2 (+W)", base + (("Wind Load", 0.4),)))
                specs.append(ComboSpec("Case 2 (-W)", base + (("Wind Load", -0.4),)))

    # ------------------------------------------------------------------
    # Case 3: (1.25D or 0.9D) + 1.5S [+ 1.0L or +/- 0.4W]
    # ------------------------------------------------------------------
    if has_s:
        for d_terms in dead_variants:
            base = d_terms + (("Snow Load", 1.5),)
            specs.append(ComboSpec("Case 3", base))
            if has_l:
                specs.append(ComboSpec("Case 3 (+L)", base + (("Live Load", 1.0),)))
            if has_w:
                specs.append(ComboSpec("Case 3 (+W)", base + (("Wind Load", 0.4),)))
                specs.append(ComboSpec("Case 3 (-W)", base + (("Wind Load", -0.4),)))

    # ------------------------------------------------------------------
    # Case 4: (1.25D or 0.9D) +/- 1.4W [+ c_L*L or + 0.5S]
    # ------------------------------------------------------------------
    if has_w:
        for d_terms in dead_variants:
            for w_sign in (1.4, -1.4):
                base = d_terms + (("Wind Load", w_sign),)
                specs.append(ComboSpec("Case 4", base))
                if has_l:
                    specs.append(ComboSpec("Case 4 (+L)", base + (("Live Load", c_l),)))
                if has_s:
                    specs.append(ComboSpec("Case 4 (+S)", base + (("Snow Load", 0.5),)))

    # ------------------------------------------------------------------
    # Case 5: 1.0D +/- 1.0E [+ c_L*L + 0.25S]
    # ------------------------------------------------------------------
    if has_e:
        companion_e: list[tuple[str, float]] = []
        if has_l:
            companion_e.append(("Live Load", c_l))
        if has_s:
            companion_e.append(("Snow Load", 0.25))

        for e_sign in (1.0, -1.0):
            base = dead_seismic + (("Seismic Load", e_sign),)
            specs.append(ComboSpec("Case 5", base))
            if companion_e:
                specs.append(
                    ComboSpec("Case 5 (+companion)", base + tuple(companion_e))
                )

    return specs


def _format_factor(value: float) -> str:
    """Format a positive load factor cleanly while preserving at least one decimal."""
    text = f"{value:.4f}".rstrip("0")
    if text.endswith("."):
        text += "0"
    return text


def _format_case_token(load_case: str) -> str:
    """Map canonical category names to compact NBCC formula symbols (D/L/S/W/E)
    and wrap custom ETABS case names containing whitespace or '+' in parentheses.
    """
    if load_case in NBCC_CATEGORY_SYMBOLS:
        return NBCC_CATEGORY_SYMBOLS[load_case]
    if any(ch in load_case for ch in (" ", "\t", "+")):
        return f"({load_case})"
    return load_case


def format_combo_name(term_cases: Sequence[tuple[str, float]]) -> str:
    """Build a human-readable NBCC combination string from `(LoadCase, factor)` pairs."""
    parts: list[str] = []
    for idx, (case_name, factor) in enumerate(term_cases):
        token = _format_case_token(case_name)
        mag = _format_factor(abs(factor))
        if idx == 0:
            sign = "-" if factor < 0 else ""
            parts.append(f"{sign}{mag}{token}")
        else:
            op = "-" if factor < 0 else "+"
            parts.append(f" {op} {mag}{token}")
    return "".join(parts)


def default_combo_definitions(
    present_categories: set[str] | frozenset[str],
    nbcc_options: NBCCOptions,
) -> list[ComboDefinition]:
    """Build editable `ComboDefinition` rows from NBCC 2020 Table 4.1.3.2.-A."""
    specs = build_nbcc_2020_specs(present_categories, nbcc_options)
    definitions: list[ComboDefinition] = []
    for spec in specs:
        definitions.append(
            ComboDefinition(
                name=format_combo_name(spec.terms),
                factors={cat: factor for cat, factor in spec.terms},
                enabled=True,
                case_id=spec.case_id,
            )
        )
    return definitions


def combo_definition_to_ordered_terms(
    combo: ComboDefinition,
) -> tuple[tuple[str, float], ...]:
    """Return `(category, factor)` terms preserving the definition's term order
    (e.g., principal load before companion loads for NBCC 2020 combinations).
    """
    return tuple(
        (cat, float(factor))
        for cat, factor in combo.factors.items()
        if factor != 0.0
    )


def resolve_active_combo_definitions(
    present_categories: set[str] | frozenset[str],
    nbcc_options: NBCCOptions,
) -> list[ComboDefinition]:
    """Return the active (enabled) combo definitions applicable to `present_categories`."""
    cats = frozenset(normalize_nbcc_category(c) for c in present_categories)
    if nbcc_options.custom_combinations is None:
        return default_combo_definitions(cats, nbcc_options)

    active: list[ComboDefinition] = []
    for combo in nbcc_options.custom_combinations:
        if not combo.enabled:
            continue
        combo.validate()
        if frozenset(combo.factors.keys()).issubset(cats):
            active.append(combo)
    return active


def _build_expanded_combo_name(
    combo_def: ComboDefinition,
    ordered_terms: Sequence[tuple[str, float]],
    chosen: Sequence[tuple[str, str, float, tuple[float, ...]]],
) -> str:
    """Build the final combo label after expanding directional/multi-step cases.

    - If `combo_def.name` matches the auto-generated NBCC formula (e.g.,
      `1.0D + 1.0E + 0.5L + 0.25S`), substitute actual load case names inline
      (`1.0D + 1.0EQx-1 + 0.5L + 0.25S`).
    - If the user gave the combination a custom name (e.g., `ULS_Wall_1`), preserve
      that custom name and append expanded case tags in parentheses when any
      load case name differs from its category name/symbol (e.g., `ULS_Wall_1 (EQx-1)`).
    """
    auto_cat_name = format_combo_name(ordered_terms)
    if combo_def.name == auto_cat_name:
        return format_combo_name([(lc, factor) for lc, _, factor, _ in chosen])

    expanded_tags = [
        lc
        for lc, cat, _, _ in chosen
        if lc != cat and lc != NBCC_CATEGORY_SYMBOLS.get(cat, cat)
    ]
    if expanded_tags:
        return f"{combo_def.name} ({', '.join(expanded_tags)})"
    return combo_def.name


def _empty_combos_df() -> pd.DataFrame:
    df = pd.DataFrame(columns=list(COMBO_OUTPUT_COLUMNS))
    for col in (*FORCE_COLUMNS, *CM_COLUMNS):
        df[col] = df[col].astype(float)
    return df


def generate_nbcc_combinations(
    cleaned_df: pd.DataFrame,
    nbcc_options: NBCCOptions,
    *,
    round_decimals: int | None = None,
) -> pd.DataFrame:
    """Compute NBCC 2020 Table 4.1.3.2.-A ULS load combinations via linear superposition.

    Args:
        cleaned_df: Output from `src.cleaner.clean_pier_forces`, containing
            `Story`, `Pier`, `LoadCase`, `Category`, `Location`, `P`, `V2`,
            `V3`, `T`, `M2`, `M3`.
        nbcc_options: NBCC 2020 options (e.g., Storage/Assembly occupancy).
        round_decimals: Optional rounding decimal places applied after full-precision
            superposition. Pass `None` to retain full floating-point precision.

    Returns:
        DataFrame with columns:
            `Story`, `Pier`, `Location`, `ComboName`, `P`, `V2`, `V3`, `T`, `M2`, `M3`.
    """
    required = ("Story", "Pier", "LoadCase", "Category", "Location", *FORCE_COLUMNS)
    missing = [c for c in required if c not in cleaned_df.columns]
    if missing:
        raise ValueError(
            f"Cleaned pier force DataFrame is missing required column(s): {missing}"
        )

    if cleaned_df.empty:
        return _empty_combos_df()

    records: list[dict[str, object]] = []

    # Preserve cleaned_df's Story, Pier, and Location ordering via sort=False.
    for (story, pier, location), group in cleaned_df.groupby(
        ["Story", "Pier", "Location"], sort=False
    ):
        # Bucket rows by canonical NBCC category while preserving row order inside each category.
        by_category: dict[str, list[tuple[str, tuple[float, ...]]]] = {}
        for row in group.itertuples(index=False):
            cat = normalize_nbcc_category(str(row.Category))
            lc = str(row.LoadCase)
            vals = (
                float(row.P),
                float(row.V2),
                float(row.V3),
                float(row.T),
                float(row.M2),
                float(row.M3),
            )
            by_category.setdefault(cat, []).append((lc, vals))

        active_defs = resolve_active_combo_definitions(
            set(by_category.keys()), nbcc_options
        )
        for combo_def in active_defs:
            ordered_terms = combo_definition_to_ordered_terms(combo_def)
            if not ordered_terms:
                continue
            # Each category in the definition may have 1 entry (e.g., aggregated D/L/S)
            # or multiple entries (e.g., WX & WY, or EQx-1, EQx-2, EQx-3, EQy).
            # Cartesian product expands each directional/stepped case separately.
            term_choices = [
                [(lc, cat, factor, vals) for lc, vals in by_category[cat]]
                for cat, factor in ordered_terms
            ]
            for chosen in product(*term_choices):
                combo_name = _build_expanded_combo_name(
                    combo_def, ordered_terms, chosen
                )
                p = v2 = v3 = t = m2 = m3 = 0.0
                for _, _, factor, vals in chosen:
                    p += factor * vals[0]
                    v2 += factor * vals[1]
                    v3 += factor * vals[2]
                    t += factor * vals[3]
                    m2 += factor * vals[4]
                    m3 += factor * vals[5]

                records.append(
                    {
                        "Story": story,
                        "Pier": pier,
                        "Location": location,
                        "ComboName": combo_name,
                        "P": p,
                        "V2": v2,
                        "V3": v3,
                        "T": t,
                        "M2": m2,
                        "M3": m3,
                    }
                )

    if not records:
        return _empty_combos_df()

    result = pd.DataFrame(records)
    result = _attach_cm_columns(result)
    result = result[list(COMBO_OUTPUT_COLUMNS)]
    if round_decimals is not None:
        result = round_force_columns(result, round_decimals)
        result = _round_cm_columns(result, round_decimals)
    return result


def envelope_columns() -> list[str]:
    """Return ordered column names for the concurrent-force Envelope Summary table."""
    return list(ENVELOPE_COLUMNS)


def compute_envelope_summary(
    combos_df: pd.DataFrame,
    *,
    round_decimals: int | None = None,
) -> pd.DataFrame:
    """Compute row-per-extreme concurrent-force envelope grouped by (Pier, Story, Location).

    For each `(Pier, Story, Location)` group, emits 10 rows (`P Min`, `P Max`,
    `V2 Min`, `V2 Max`, `V3 Min`, `V3 Max`, `M2 Min`, `M2 Max`, `M3 Min`, `M3 Max`),
    carrying all concurrent force components (`P`, `T`, `V2`, `M3`, `V3`, `M2`) from
    the governing load combination row.

    Args:
        combos_df: Output from `generate_nbcc_combinations`.
        round_decimals: Optional decimal places for numeric force columns.

    Returns:
        DataFrame with columns:
            `Pier`, `Story`, `Combo`, `Extreme`, `Location`, `P`, `T`, `V2`, `M3`, `V3`, `M2`.
    """
    cols = envelope_columns()
    required = ("Story", "Pier", "Location", "ComboName", *FORCE_COLUMNS)
    missing = [c for c in required if c not in combos_df.columns]
    if missing:
        raise ValueError(
            f"Combinations DataFrame is missing required column(s): {missing}"
        )

    if combos_df.empty:
        empty = pd.DataFrame(columns=cols)
        for col in (*_ENVELOPE_FORCE_COLUMNS, *CM_COLUMNS):
            empty[col] = empty[col].astype(float)
        return empty

    work_df = combos_df
    if "Cm3" not in work_df.columns or "Cm2" not in work_df.columns:
        work_df = _attach_cm_columns(work_df)

    rows: list[dict[str, object]] = []
    cm_dec = max(2, round_decimals) if round_decimals is not None else None

    for (pier, story, location), group in work_df.groupby(
        ["Pier", "Story", "Location"], sort=False
    ):
        for comp in ENVELOPE_EXTREME_COMPONENTS:
            min_idx = group[comp].idxmin()
            max_idx = group[comp].idxmax()

            for extreme_label, idx in ((f"{comp} Min", min_idx), (f"{comp} Max", max_idx)):
                winning = group.loc[idx]
                record: dict[str, object] = {
                    "Pier": pier,
                    "Story": story,
                    "Combo": str(winning["ComboName"]),
                    "Extreme": extreme_label,
                    "Location": location,
                }
                for force_col in _ENVELOPE_FORCE_COLUMNS:
                    val = float(winning[force_col])
                    if round_decimals is not None:
                        val = round(val, round_decimals)
                        if val == 0.0:
                            val = 0.0
                    record[force_col] = val
                for cm_col in CM_COLUMNS:
                    cm_val = float(winning[cm_col])
                    if cm_dec is not None:
                        cm_val = round(cm_val, cm_dec)
                    record[cm_col] = cm_val
                rows.append(record)

    return pd.DataFrame(rows, columns=cols)


def run_nbcc_analysis(
    cleaned_df: pd.DataFrame,
    cleaning_options: CleaningOptions,
    nbcc_options: NBCCOptions,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Compute NBCC 2020 combinations (including paired Top/Bottom Cm3 & Cm2) at full
    precision, apply the Station Filter, compute the governing envelope, and round
    displayed outputs to `cleaning_options.round_decimals`.
    """
    cleaning_options.validate()
    combos_full = generate_nbcc_combinations(
        cleaned_df, nbcc_options, round_decimals=None
    )

    # Apply Station Filter AFTER Cm3/Cm2 have been computed from paired Top/Bottom moments.
    if cleaning_options.station_filter != "All" and not combos_full.empty:
        combos_filtered = combos_full[
            combos_full["Location"] == cleaning_options.station_filter
        ].reset_index(drop=True)
    else:
        combos_filtered = combos_full

    envelope = compute_envelope_summary(
        combos_filtered,
        round_decimals=cleaning_options.round_decimals,
    )
    combos_rounded = round_force_columns(
        combos_filtered, cleaning_options.round_decimals
    )
    combos_rounded = _round_cm_columns(
        combos_rounded, cleaning_options.round_decimals
    )
    return combos_rounded, envelope
