"""Unit tests for src.nbcc_combos (NBCC 2020 Table 4.1.3.2.-A)."""

from __future__ import annotations

import pandas as pd
import pytest

from src.config import CleaningOptions, ComboDefinition, NBCCOptions
from src.nbcc_combos import (
    compute_cm_factor,
    compute_envelope_summary,
    default_combo_definitions,
    generate_nbcc_combinations,
    run_nbcc_analysis,
)


def _full_single_station_cleaned_df() -> pd.DataFrame:
    """Single (Story, Pier, Location) containing D, L, S, two Wind cases, and multi-step E."""
    return pd.DataFrame(
        [
            {"Story": "L2", "Pier": "P1", "LoadCase": "D", "Category": "D", "Location": "Bottom", "P": -1000.0, "V2": 10.0, "V3": 2.0, "T": 1.0, "M2": 5.0, "M3": 20.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "L", "Category": "L", "Location": "Bottom", "P": -400.0, "V2": 4.0, "V3": 1.0, "T": 0.5, "M2": 2.0, "M3": 8.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "S", "Category": "S", "Location": "Bottom", "P": -200.0, "V2": 2.0, "V3": 0.5, "T": 0.2, "M2": 1.0, "M3": 4.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "WX", "Category": "W", "Location": "Bottom", "P": 100.0, "V2": 250.0, "V3": 10.0, "T": 15.0, "M2": 30.0, "M3": 1000.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "WY", "Category": "W", "Location": "Bottom", "P": -50.0, "V2": 60.0, "V3": 200.0, "T": 25.0, "M2": 400.0, "M3": 200.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx-1", "Category": "E", "Location": "Bottom", "P": 300.0, "V2": 600.0, "V3": 20.0, "T": 50.0, "M2": 100.0, "M3": 3000.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx-2", "Category": "E", "Location": "Bottom", "P": 320.0, "V2": 640.0, "V3": 25.0, "T": 55.0, "M2": 110.0, "M3": 3200.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQy", "Category": "E", "Location": "Bottom", "P": 150.0, "V2": 100.0, "V3": 500.0, "T": 80.0, "M2": 900.0, "M3": 500.0},
        ]
    )


def test_nbcc_2020_standard_occupancy_exact_hand_calculations() -> None:
    df = _full_single_station_cleaned_df()
    combos = generate_nbcc_combinations(
        df, NBCCOptions(storage_assembly_occupancy=False)
    )
    by_name = combos.set_index("ComboName")

    # Expected count:
    # Case 1: 1
    # Case 2: 2 base + 2 (+S) + 2*2*2 (+/-0.4W for WX, WY) = 12
    # Case 3: 2 base + 2 (+L) + 2*2*2 (+/-0.4W for WX, WY) = 12
    # Case 4: 2 dead * 2 W cases * 2 signs * 3 (base, +0.5L, +0.5S) = 24
    # Case 5: 3 E entries (EQx-1, EQx-2, EQy) * 2 signs * 2 (base, +0.5L+0.25S) = 12
    assert len(combos) == 1 + 12 + 12 + 24 + 12

    # Case 1: 1.4D
    row = by_name.loc["1.4D"]
    assert row["P"] == pytest.approx(-1400.0)
    assert row["V2"] == pytest.approx(14.0)
    assert row["T"] == pytest.approx(1.4)
    assert row["M3"] == pytest.approx(28.0)

    # Case 2: 1.25D + 1.5L + 1.0S
    row = by_name.loc["1.25D + 1.5L + 1.0S"]
    assert row["P"] == pytest.approx(1.25 * (-1000) + 1.5 * (-400) + 1.0 * (-200))  # -2050
    assert row["M3"] == pytest.approx(1.25 * 20 + 1.5 * 8 + 1.0 * 4)  # 41

    # Case 2 with wind reversal: 0.9D + 1.5L - 0.4WX
    row = by_name.loc["0.9D + 1.5L - 0.4WX"]
    assert row["P"] == pytest.approx(0.9 * (-1000) + 1.5 * (-400) - 0.4 * 100)  # -1540
    assert row["V2"] == pytest.approx(0.9 * 10 + 1.5 * 4 - 0.4 * 250)  # -85

    # Case 3: 1.25D + 1.5S + 1.0L (Case 3 companion L is ALWAYS 1.0L per NBCC 2020)
    row = by_name.loc["1.25D + 1.5S + 1.0L"]
    assert row["P"] == pytest.approx(1.25 * (-1000) + 1.5 * (-200) + 1.0 * (-400))  # -1950

    # Case 4: 0.9D + 1.4WX + 0.5L (standard occupancy -> 0.5L)
    row = by_name.loc["0.9D + 1.4WX + 0.5L"]
    assert row["P"] == pytest.approx(0.9 * (-1000) + 1.4 * 100 + 0.5 * (-400))  # -960
    assert row["V2"] == pytest.approx(0.9 * 10 + 1.4 * 250 + 0.5 * 4)  # 361

    # Case 4: 1.25D - 1.4WY + 0.5S
    row = by_name.loc["1.25D - 1.4WY + 0.5S"]
    assert row["V3"] == pytest.approx(1.25 * 2 - 1.4 * 200 + 0.5 * 0.5)  # -277.25

    # Case 5: 1.0D + 1.0EQx-2 + 0.5L + 0.25S (standard occupancy)
    row = by_name.loc["1.0D + 1.0EQx-2 + 0.5L + 0.25S"]
    assert row["P"] == pytest.approx(-1000 + 320 + 0.5 * (-400) + 0.25 * (-200))  # -930
    assert row["M3"] == pytest.approx(20 + 3200 + 0.5 * 8 + 0.25 * 4)  # 3225

    # Case 5 without companion: 1.0D - 1.0EQx-2
    row = by_name.loc["1.0D - 1.0EQx-2"]
    assert row["P"] == pytest.approx(-1000 - 320)  # -1320
    assert row["M3"] == pytest.approx(20 - 3200)  # -3180


def test_nbcc_2020_storage_assembly_occupancy_changes_case_4_and_5_only() -> None:
    df = _full_single_station_cleaned_df()
    combos = generate_nbcc_combinations(
        df, NBCCOptions(storage_assembly_occupancy=True)
    )
    by_name = combos.set_index("ComboName")

    # Case 3 remains 1.0L.
    assert "1.25D + 1.5S + 1.0L" in by_name.index

    # Case 4 companion L increases from 0.5L to 1.0L.
    assert "1.25D + 1.4WX + 1.0L" in by_name.index
    assert "1.25D + 1.4WX + 0.5L" not in by_name.index
    assert by_name.loc["1.25D + 1.4WX + 1.0L", "P"] == pytest.approx(
        1.25 * (-1000) + 1.4 * 100 + 1.0 * (-400)
    )

    # Case 5 companion L increases from 0.5L to 1.0L.
    assert "1.0D + 1.0EQx-1 + 1.0L + 0.25S" in by_name.index
    assert "1.0D + 1.0EQx-1 + 0.5L + 0.25S" not in by_name.index
    assert by_name.loc["1.0D + 1.0EQx-1 + 1.0L + 0.25S", "P"] == pytest.approx(
        -1000 + 300 + 1.0 * (-400) + 0.25 * (-200)
    )


def test_subset_of_categories_generates_only_valid_combos() -> None:
    # Only D and E mapped (no L, S, W).
    df = pd.DataFrame(
        [
            {"Story": "L1", "Pier": "P1", "LoadCase": "D", "Category": "D", "Location": "Bottom", "P": -500.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
            {"Story": "L1", "Pier": "P1", "LoadCase": "EQx", "Category": "E", "Location": "Bottom", "P": 100.0, "V2": 200.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 800.0},
        ]
    )
    combos = generate_nbcc_combinations(df, NBCCOptions())
    assert list(combos["ComboName"]) == ["1.4D", "1.0D + 1.0EQx", "1.0D - 1.0EQx"]


def test_envelope_summary_grouping_and_governing_combos() -> None:
    df = _full_single_station_cleaned_df()
    # Add a second station (Top) at Roof/P2 to test ordering & (Pier, Story, Location) grouping.
    top_rows = pd.DataFrame(
        [
            {"Story": "Roof", "Pier": "P2", "LoadCase": "D", "Category": "D", "Location": "Top", "P": -200.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
            {"Story": "Roof", "Pier": "P2", "LoadCase": "L", "Category": "L", "Location": "Top", "P": -100.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
        ]
    )
    combined = pd.concat([top_rows, df], ignore_index=True)

    combos, env = run_nbcc_analysis(
        combined,
        CleaningOptions(station_filter="All", round_decimals=2),
        NBCCOptions(storage_assembly_occupancy=False),
    )

    # Exact column order matching user screenshot + Cm3/Cm2.
    assert list(env.columns) == [
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
    ]

    # 2 (Pier, Story, Location) groups x 10 extremes (P, V2, V3, M2, M3 Min/Max) = 20 rows.
    assert len(env) == 20
    expected_extremes = [
        "P Min",
        "P Max",
        "V2 Min",
        "V2 Max",
        "V3 Min",
        "V3 Max",
        "M2 Min",
        "M2 Max",
        "M3 Min",
        "M3 Max",
    ]

    # Group 1: P2 / Roof / Top
    p2_env = env[(env["Pier"] == "P2") & (env["Story"] == "Roof") & (env["Location"] == "Top")]
    assert list(p2_env["Extreme"]) == expected_extremes

    # Group 2: P1 / L2 / Bottom — verify governing combos AND concurrent forces!
    p1_env = (
        env[(env["Pier"] == "P1") & (env["Story"] == "L2") & (env["Location"] == "Bottom")]
        .set_index("Extreme")
    )
    assert list(p1_env.index) == expected_extremes

    # P Min (max compression) = 1.25D + 1.5L + 1.0S -> P = -2050.0, concurrent M3 = 41.0
    assert p1_env.loc["P Min", "Combo"] == "1.25D + 1.5L + 1.0S"
    assert p1_env.loc["P Min", "P"] == pytest.approx(-2050.0)
    assert p1_env.loc["P Min", "M3"] == pytest.approx(41.0)

    # P Max (least compression) = 1.0D + 1.0EQx-2 -> P = -680.0, concurrent V2 = 650.0, M3 = 3220.0
    assert p1_env.loc["P Max", "Combo"] == "1.0D + 1.0EQx-2"
    assert p1_env.loc["P Max", "P"] == pytest.approx(-680.0)
    assert p1_env.loc["P Max", "V2"] == pytest.approx(650.0)
    assert p1_env.loc["P Max", "M3"] == pytest.approx(3220.0)

    # M3 Max = 1.0D + 1.0EQx-2 + 0.5L + 0.25S -> M3 = 3225.0, concurrent P = -930.0
    assert p1_env.loc["M3 Max", "Combo"] == "1.0D + 1.0EQx-2 + 0.5L + 0.25S"
    assert p1_env.loc["M3 Max", "M3"] == pytest.approx(3225.0)
    assert p1_env.loc["M3 Max", "P"] == pytest.approx(-930.0)

    # M3 Min = 1.0D - 1.0EQx-2 -> M3 = -3180.0, concurrent P = -1320.0
    assert p1_env.loc["M3 Min", "Combo"] == "1.0D - 1.0EQx-2"
    assert p1_env.loc["M3 Min", "M3"] == pytest.approx(-3180.0)
    assert p1_env.loc["M3 Min", "P"] == pytest.approx(-1320.0)

    # Single-station input still retains Location column.
    env_single = compute_envelope_summary(
        combos[combos["Location"] == "Bottom"],
        round_decimals=2,
    )
    assert list(env_single.columns) == list(env.columns)
    assert len(env_single) == 10


def test_cm_factor_csa_a23_3_single_double_curvature_and_edge_cases() -> None:
    # Single curvature (same sign): Top=2000, Bot=4000 -> M1/M2 = +0.5 -> 0.6 + 0.4(0.5) = 0.8
    assert compute_cm_factor(2000.0, 4000.0) == pytest.approx(0.8)
    # Single curvature (both negative): Top=-2000, Bot=-4000 -> +0.5 -> 0.8
    assert compute_cm_factor(-2000.0, -4000.0) == pytest.approx(0.8)
    # Double curvature (opposite sign): Top=-500, Bot=2000 -> M1/M2 = -0.25 -> 0.6 - 0.1 = 0.5
    assert compute_cm_factor(-500.0, 2000.0) == pytest.approx(0.5)
    # Pure double curvature hits 0.4 floor: Top=-1000, Bot=1000 -> 0.6 - 0.4 = 0.2 -> clamped to 0.4
    assert compute_cm_factor(-1000.0, 1000.0) == pytest.approx(0.4)
    # Larger moment at Top than Bottom: |M_top|=4000, |M_bot|=1000 -> M1/M2 = +0.25 -> 0.7
    assert compute_cm_factor(4000.0, 1000.0) == pytest.approx(0.7)
    # Zero moment at one end: M1=0 -> 0.6
    assert compute_cm_factor(0.0, 1000.0) == pytest.approx(0.6)
    # Both zero or missing end -> conservative default 1.0
    assert compute_cm_factor(0.0, 0.0) == pytest.approx(1.0)
    assert compute_cm_factor(None, 1000.0) == pytest.approx(1.0)


def test_bottom_only_station_filter_computes_cm3_and_cm2_from_paired_top_and_bottom() -> None:
    # Single story/pier with BOTH Top & Bottom rows; Station Filter = Bottom.
    cleaned = pd.DataFrame(
        [
            # Dead (zero moments)
            {"Story": "L2", "Pier": "P1", "LoadCase": "D", "Category": "D", "Location": "Top", "P": -500.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "D", "Category": "D", "Location": "Bottom", "P": -550.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
            # EQx: M3 single curvature (+1000 top, +2000 bot -> Cm3 = 0.80)
            #      M2 double curvature (-250 top, +1000 bot -> Cm2 = 0.50)
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx", "Category": "E", "Location": "Top", "P": 0.0, "V2": 200.0, "V3": 50.0, "T": 0.0, "M2": -250.0, "M3": 1000.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx", "Category": "E", "Location": "Bottom", "P": 0.0, "V2": 200.0, "V3": 50.0, "T": 0.0, "M2": 1000.0, "M3": 2000.0},
        ]
    )

    combos, env = run_nbcc_analysis(
        cleaned,
        CleaningOptions(station_filter="Bottom", round_decimals=0),
        NBCCOptions(),
    )

    # Output contains ONLY Bottom rows, yet Cm3 & Cm2 were calculated from Top & Bottom!
    assert set(combos["Location"]) == {"Bottom"}
    assert set(env["Location"]) == {"Bottom"}

    by_combo = combos.set_index("ComboName")
    assert by_combo.loc["1.4D", "Cm3"] == pytest.approx(1.0)
    assert by_combo.loc["1.4D", "Cm2"] == pytest.approx(1.0)
    # Even with round_decimals=0, Cm3 and Cm2 preserve 2 decimal places!
    assert by_combo.loc["1.0D + 1.0EQx", "Cm3"] == pytest.approx(0.80)
    assert by_combo.loc["1.0D + 1.0EQx", "Cm2"] == pytest.approx(0.50)
    # Reversed sign (-EQx) preserves same curvature ratios!
    assert by_combo.loc["1.0D - 1.0EQx", "Cm3"] == pytest.approx(0.80)
    assert by_combo.loc["1.0D - 1.0EQx", "Cm2"] == pytest.approx(0.50)

    by_ext = env.set_index("Extreme")
    assert by_ext.loc["M3 Max", "M3"] == pytest.approx(2000.0)
    assert by_ext.loc["M3 Max", "Cm3"] == pytest.approx(0.80)
    assert by_ext.loc["M3 Max", "Cm2"] == pytest.approx(0.50)


def test_custom_combinations_respect_enabled_and_expand_directional_cases() -> None:
    df = _full_single_station_cleaned_df()
    custom = [
        ComboDefinition(name="SLS_Gravity", factors={"D": 1.0, "L": 1.0}, enabled=True),
        ComboDefinition(name="Disabled_Combo", factors={"D": 1.4}, enabled=False),
        ComboDefinition(name="Custom_EQ", factors={"D": 1.0, "E": 0.75}, enabled=True),
    ]
    combos = generate_nbcc_combinations(
        df, NBCCOptions(custom_combinations=custom)
    )

    names = list(combos["ComboName"])
    assert "Disabled_Combo" not in names
    assert "SLS_Gravity" in names
    # Custom_EQ expands across EQx-1, EQx-2, and EQy cleanly!
    assert "Custom_EQ (EQx-1)" in names
    assert "Custom_EQ (EQx-2)" in names
    assert "Custom_EQ (EQy)" in names
    assert len(combos) == 4

    by_name = combos.set_index("ComboName")
    assert by_name.loc["SLS_Gravity", "P"] == pytest.approx(-1400.0)
    assert by_name.loc["Custom_EQ (EQx-2)", "M3"] == pytest.approx(20.0 + 0.75 * 3200.0)

    # Default definitions helper matches NBCC 2020 count.
    defaults = default_combo_definitions({"D", "L"}, NBCCOptions())
    assert [d.name for d in defaults] == ["1.4D", "1.25D + 1.5L", "0.9D + 1.5L"]
