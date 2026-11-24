"""Unit tests for src.cleaner."""

from __future__ import annotations

import pandas as pd
import pytest

from src.cleaner import clean_pier_forces, round_force_columns
from src.config import CleaningOptions


def _sample_raw_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            # Gravity sub-cases (DL_Self + SDL -> D)
            {"Story": "Roof", "Pier": "P2", "LoadCase": "DL_Self", "Location": "Top", "P": -100.123, "V2": 2.0, "V3": 1.0, "T": 0.5, "M2": 3.0, "M3": 10.0},
            {"Story": "Roof", "Pier": "P2", "LoadCase": "SDL", "Location": "Top", "P": -50.456, "V2": 1.0, "V3": 0.5, "T": 0.25, "M2": 1.5, "M3": 5.0},
            {"Story": "Roof", "Pier": "P2", "LoadCase": "DL_Self", "Location": "Bottom", "P": -110.0, "V2": 2.0, "V3": 1.0, "T": 0.5, "M2": 4.0, "M3": 12.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "DL_Self", "Location": "Bottom", "P": -200.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
            # Multi-step seismic EQx (3 accidental eccentricity steps)
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx", "Location": "Bottom", "P": -99.5, "V2": -172.75, "V3": -179.8, "T": 222.9, "M2": -594.5, "M3": -745.8},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx", "Location": "Bottom", "P": -124.6, "V2": -164.02, "V3": -149.3, "T": 188.5, "M2": -489.9, "M3": -680.3},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx", "Location": "Bottom", "P": -74.4, "V2": -181.48, "V3": -210.2, "T": 257.2, "M2": -699.1, "M3": -811.2},
            # Second directional seismic EQy
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQy", "Location": "Bottom", "P": 50.0, "V2": 300.0, "V3": 20.0, "T": -10.0, "M2": 40.0, "M3": 1200.0},
            # Near-zero noise row
            {"Story": "L2", "Pier": "P1", "LoadCase": "LIVE", "Location": "Bottom", "P": 1e-6, "V2": -2e-6, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
            # Unmapped case (must be ignored)
            {"Story": "L2", "Pier": "P1", "LoadCase": "MODAL", "Location": "Bottom", "P": 999.0, "V2": 999.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 0.0},
        ]
    )


def test_clean_aggregates_gravity_and_preserves_multistep_seismic() -> None:
    raw = _sample_raw_df()
    mapping = {"DL_Self": "D", "SDL": "D", "LIVE": "L", "EQx": "E", "EQy": "E"}
    opts = CleaningOptions(
        station_filter="All", aggregate_subcases=True, drop_near_zero=False
    )

    cleaned = clean_pier_forces(
        raw,
        mapping,
        opts,
        story_order=["Roof", "L2"],
        pier_order=["P1", "P2"],
    )

    # Unmapped MODAL should not appear.
    assert "MODAL" not in set(cleaned["LoadCase"])

    # Gravity sub-cases DL_Self + SDL at Roof/P2/Top summed at full precision.
    roof_d = cleaned[
        (cleaned["Story"] == "Roof")
        & (cleaned["Pier"] == "P2")
        & (cleaned["Location"] == "Top")
        & (cleaned["Category"] == "Dead Load")
    ]
    assert len(roof_d) == 1
    assert roof_d.iloc[0]["LoadCase"] == "Dead Load"
    assert roof_d.iloc[0]["P"] == pytest.approx(-150.579)
    assert roof_d.iloc[0]["T"] == pytest.approx(0.75)

    # Multi-step EQx must NOT be summed (fixes PyCSI 3x seismic summing bug).
    eqx_rows = cleaned[cleaned["LoadCase"].str.startswith("EQx")]
    assert list(eqx_rows["LoadCase"]) == ["EQx-1", "EQx-2", "EQx-3"]
    assert list(eqx_rows["V2"]) == pytest.approx([-172.75, -164.02, -181.48])

    # EQy stays distinct from EQx even though both map to category E.
    eqy_rows = cleaned[cleaned["LoadCase"] == "EQy"]
    assert len(eqy_rows) == 1
    assert eqy_rows.iloc[0]["V2"] == pytest.approx(300.0)


def test_clean_filters_story_pier_station_and_preserves_order() -> None:
    raw = _sample_raw_df()
    mapping = {"DL_Self": "D", "SDL": "D", "EQx": "E", "EQy": "E"}
    opts = CleaningOptions(
        selected_piers=["P2"],
        selected_stories=["Roof"],
        station_filter="Bottom",
    )

    cleaned = clean_pier_forces(raw, mapping, opts)
    assert len(cleaned) == 1
    assert cleaned.iloc[0]["Story"] == "Roof"
    assert cleaned.iloc[0]["Pier"] == "P2"
    assert cleaned.iloc[0]["Location"] == "Bottom"


def test_clean_near_zero_and_component_zeroing_and_duplicates() -> None:
    raw = _sample_raw_df()
    # Add an exact duplicate of the EQy row.
    raw = pd.concat([raw, raw[raw["LoadCase"] == "EQy"]], ignore_index=True)

    mapping = {"DL_Self": "D", "SDL": "D", "LIVE": "L", "EQx": "E", "EQy": "E"}
    opts = CleaningOptions(
        selected_stories=["L2"],
        selected_piers=["P1"],
        drop_near_zero=True,
        near_zero_threshold=1e-4,
        remove_duplicates=True,
        zero_lateral_axial=True,
        zero_gravity_shear_moment=True,
    )

    cleaned = clean_pier_forces(raw, mapping, opts)

    # Near-zero LIVE row dropped.
    assert "Live Load" not in set(cleaned["Category"])

    # Exact duplicate of EQy removed -> single EQy row (not EQy-1 / EQy-2).
    eqy_rows = cleaned[cleaned["LoadCase"].str.startswith("EQy")]
    assert list(eqy_rows["LoadCase"]) == ["EQy"]

    # Distinct multi-step EQx rows were NOT dropped by remove_duplicates.
    eqx_rows = cleaned[cleaned["LoadCase"].str.startswith("EQx")]
    assert list(eqx_rows["LoadCase"]) == ["EQx-1", "EQx-2", "EQx-3"]

    # Lateral axial P zeroed.
    assert (cleaned.loc[cleaned["Category"] == "Seismic Load", "P"] == 0.0).all()

    # Gravity shear & moment zeroed while preserving P.
    grav = cleaned[cleaned["Category"] == "Dead Load"]
    assert grav.iloc[0]["P"] == pytest.approx(-200.0)
    for col in ("V2", "V3", "T", "M2", "M3"):
        assert grav.iloc[0][col] == 0.0


def test_clean_without_subcase_aggregation_keeps_gravity_cases_distinct() -> None:
    raw = _sample_raw_df()
    mapping = {"DL_Self": "D", "SDL": "D"}
    opts = CleaningOptions(
        selected_stories=["Roof"],
        selected_piers=["P2"],
        station_filter="Top",
        aggregate_subcases=False,
    )

    cleaned = clean_pier_forces(raw, mapping, opts)
    assert list(cleaned["LoadCase"]) == ["DL_Self", "SDL"]


def test_round_force_columns_and_missing_column_error() -> None:
    df = pd.DataFrame([{"P": -0.0001, "V2": 12.3456, "T": 1.999}])
    rounded = round_force_columns(df, 2)
    assert rounded.iloc[0]["P"] == 0.0
    assert str(rounded.iloc[0]["P"]) == "0.0"  # not -0.0
    assert rounded.iloc[0]["V2"] == 12.35
    assert rounded.iloc[0]["T"] == 2.0

    with pytest.raises(ValueError, match="missing required column"):
        clean_pier_forces(
            pd.DataFrame([{"Story": "L1"}]),
            {"DL": "D"},
            CleaningOptions(),
        )
