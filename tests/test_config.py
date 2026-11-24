"""Unit tests for src.config."""

from __future__ import annotations

from pathlib import Path
import pytest

from src.config import (
    AppConfig,
    CleaningOptions,
    ComboDefinition,
    NBCCOptions,
    load_config,
    reconcile_config_with_model,
    save_config,
)


def test_roundtrip_save_and_load_config(tmp_path: Path) -> None:
    cfg = AppConfig(
        units="kN, m, C",
        case_mappings={
            "DL_Self": "Dead Load",
            "SDL": "Dead Load",
            "LIVE": "Live Load",
            "SNOW": "Snow Load",
            "WX": "Wind Load",
            "EQX": "Seismic Load",
            "EQY": "Seismic Load",
        },
        cleaning=CleaningOptions(
            selected_piers=["P1", "P2"],
            selected_stories=["Roof", "L3", "L2"],
            station_filter="Bottom",
            round_decimals=3,
            drop_near_zero=True,
            near_zero_threshold=1e-3,
            remove_duplicates=True,
            aggregate_subcases=True,
            zero_lateral_axial=True,
            zero_gravity_shear_moment=False,
        ),
        nbcc=NBCCOptions(storage_assembly_occupancy=True),
    )

    config_file = tmp_path / "preset.json"
    save_config(cfg, config_file)
    loaded = load_config(config_file)

    assert loaded == cfg


def test_load_config_validation_errors(tmp_path: Path) -> None:
    bad_json = tmp_path / "bad.json"
    bad_json.write_text("{not valid json", encoding="utf-8")
    with pytest.raises(ValueError, match="Invalid JSON"):
        load_config(bad_json)

    bad_cat = AppConfig(case_mappings={"DL": "INVALID"})
    with pytest.raises(ValueError, match="Invalid NBCC category"):
        bad_cat.validate()

    bad_station = AppConfig(cleaning=CleaningOptions(station_filter="Middle"))
    with pytest.raises(ValueError, match="Invalid station_filter"):
        bad_station.validate()

    bad_units = AppConfig(units="tons, furlongs, K")
    with pytest.raises(ValueError, match="Invalid units"):
        bad_units.validate()

    bad_decimals = AppConfig(cleaning=CleaningOptions(round_decimals=-1))
    with pytest.raises(ValueError, match="round_decimals"):
        bad_decimals.validate()


def test_reconcile_config_partial_missing_items() -> None:
    saved = AppConfig(
        case_mappings={
            "DL": "Dead Load",
            "SDL": "Dead Load",
            "OLD_WIND": "Wind Load",
            "EQX": "Seismic Load",
        },
        cleaning=CleaningOptions(
            selected_piers=["P1", "P2", "P99"],
            selected_stories=["Roof", "L3", "Basement"],
        ),
    )

    result = reconcile_config_with_model(
        saved,
        available_cases=["DL", "SDL", "EQX", "LIVE"],
        available_piers=["P1", "P2", "P3"],
        available_stories=["Roof", "L3", "L2"],
    )

    assert result.has_warnings
    assert result.missing_cases == ["OLD_WIND"]
    assert result.missing_piers == ["P99"]
    assert result.missing_stories == ["Basement"]
    assert result.config.case_mappings == {
        "DL": "Dead Load",
        "SDL": "Dead Load",
        "EQX": "Seismic Load",
    }
    assert result.config.cleaning.selected_piers == ["P1", "P2"]
    assert result.config.cleaning.selected_stories == ["Roof", "L3"]


def test_reconcile_config_all_saved_filters_missing_resets_to_all() -> None:
    saved = AppConfig(
        case_mappings={"DL": "D"},
        cleaning=CleaningOptions(
            selected_piers=["OLD_P1"],
            selected_stories=["OLD_S1"],
        ),
    )

    result = reconcile_config_with_model(
        saved,
        available_cases=["DL"],
        available_piers=["P1", "P2"],
        available_stories=["L2", "L1"],
    )

    assert result.has_warnings
    assert result.config.cleaning.selected_piers is None
    assert result.config.cleaning.selected_stories is None
    assert any("pier filter reset to All" in w for w in result.warnings)
    assert any("story filter reset to All" in w for w in result.warnings)


def test_custom_combinations_roundtrip_and_validation(tmp_path: Path) -> None:
    cfg = AppConfig(
        case_mappings={"DL": "Dead Load", "LL": "Live Load", "EQx": "Seismic Load"},
        nbcc=NBCCOptions(
            storage_assembly_occupancy=False,
            custom_combinations=[
                ComboDefinition(name="1.4D", factors={"Dead Load": 1.4}, enabled=True, case_id="Case 1"),
                ComboDefinition(name="SLS_Check", factors={"Dead Load": 1.0, "Live Load": 1.0}, enabled=True),
                ComboDefinition(name="Disabled_EQ", factors={"Dead Load": 1.0, "Seismic Load": 1.0}, enabled=False),
            ],
        ),
    )
    path = tmp_path / "custom_combos.json"
    save_config(cfg, path)
    loaded = load_config(path)
    assert loaded == cfg

    # Legacy shorthand JSON presets (D/L/S/W/E) normalize cleanly to full names.
    legacy_path = tmp_path / "legacy.json"
    legacy_path.write_text(
        '{"case_mappings": {"DL": "D", "LL": "L", "SL": "S", "WX": "W", "EX": "E"}}',
        encoding="utf-8",
    )
    legacy_loaded = load_config(legacy_path)
    assert legacy_loaded.case_mappings == {
        "DL": "Dead Load",
        "LL": "Live Load",
        "SL": "Snow Load",
        "WX": "Wind Load",
        "EX": "Seismic Load",
    }

    # Validation rejects empty name or all-zero factors.
    with pytest.raises(ValueError, match="non-empty"):
        ComboDefinition(name="  ", factors={"Dead Load": 1.0}).validate()
    with pytest.raises(ValueError, match="non-zero"):
        ComboDefinition(name="Empty", factors={"Dead Load": 0.0}).validate()
    with pytest.raises(ValueError, match="Invalid NBCC category"):
        ComboDefinition(name="BadCat", factors={"X": 1.0}).validate()
