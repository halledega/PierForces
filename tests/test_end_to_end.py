"""End-to-end pipeline and UI smoke tests (no live ETABS instance required)."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk

import openpyxl
import pandas as pd
import pytest

from src.cleaner import clean_pier_forces
from src.config import AppConfig, CleaningOptions, ComboDefinition, NBCCOptions
from src.etabs_client import ETABSModelInfo
from src.exporter import ExportMetadata, export_to_excel
from src.nbcc_combos import run_nbcc_analysis
from src.ui import PierForcesApp


def test_full_pipeline_cleaning_nbcc_envelope_excel(tmp_path: Path) -> None:
    raw = pd.DataFrame(
        [
            {"Story": "Roof", "Pier": "P1", "LoadCase": "DL_SW", "Location": "Bottom", "P": -300.0, "V2": 2.0, "V3": 0.5, "T": 0.1, "M2": 1.0, "M3": 5.0},
            {"Story": "Roof", "Pier": "P1", "LoadCase": "SDL", "Location": "Bottom", "P": -100.0, "V2": 1.0, "V3": 0.2, "T": 0.1, "M2": 0.5, "M3": 2.0},
            {"Story": "Roof", "Pier": "P1", "LoadCase": "LIVE", "Location": "Bottom", "P": -150.0, "V2": 1.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 3.0},
            {"Story": "Roof", "Pier": "P1", "LoadCase": "SNOW", "Location": "Bottom", "P": -80.0, "V2": 0.5, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 1.0},
            # 3-step seismic EQx
            {"Story": "Roof", "Pier": "P1", "LoadCase": "EQx", "Location": "Bottom", "P": 60.0, "V2": 210.0, "V3": 15.0, "T": 25.0, "M2": 40.0, "M3": 900.0},
            {"Story": "Roof", "Pier": "P1", "LoadCase": "EQx", "Location": "Bottom", "P": 75.0, "V2": 240.0, "V3": 18.0, "T": 30.0, "M2": 48.0, "M3": 1050.0},
            {"Story": "Roof", "Pier": "P1", "LoadCase": "EQx", "Location": "Bottom", "P": 45.0, "V2": 180.0, "V3": 12.0, "T": 20.0, "M2": 32.0, "M3": 750.0},
        ]
    )

    cfg = AppConfig(
        units="kN, m, C",
        case_mappings={
            "DL_SW": "D",
            "SDL": "D",
            "LIVE": "L",
            "SNOW": "S",
            "EQx": "E",
        },
        cleaning=CleaningOptions(
            station_filter="Bottom",
            round_decimals=2,
            aggregate_subcases=True,
        ),
        nbcc=NBCCOptions(storage_assembly_occupancy=False),
    )

    cleaned = clean_pier_forces(
        raw,
        cfg.case_mappings,
        cfg.cleaning,
        story_order=["Roof"],
        pier_order=["P1"],
    )
    combos, envelope = run_nbcc_analysis(cleaned, cfg.cleaning, cfg.nbcc)

    # Verify 10 extreme rows for (P1, Roof, Bottom) and that EQx-2 governs V2 Max (not summed 3x).
    assert len(envelope) == 10
    by_ext = envelope.set_index("Extreme")
    assert by_ext.loc["V2 Max", "Combo"] == "1.0D + 1.0EQx-2 + 0.5L + 0.25S"
    assert by_ext.loc["V2 Max", "V2"] == pytest.approx(
        round(3.0 + 240.0 + 0.5 * 1.0 + 0.25 * 0.5, 2)
    )
    assert by_ext.loc["P Min", "Combo"] == "1.25D + 1.5L + 1.0S"
    assert by_ext.loc["P Min", "P"] == pytest.approx(
        1.25 * (-400) + 1.5 * (-150) + 1.0 * (-80)
    )

    xlsx = tmp_path / "e2e.xlsx"
    export_to_excel(
        envelope,
        combos,
        cfg,
        ExportMetadata(
            model_filename="E2E.EDB",
            units="kN, m, C",
            raw_row_count=len(raw),
            combo_row_count=len(combos),
            envelope_row_count=len(envelope),
        ),
        xlsx,
    )
    wb = openpyxl.load_workbook(xlsx)
    assert wb.sheetnames == ["Envelope Summary", "Load Combinations", "Metadata"]


def test_ui_smoke_cache_recalculate_and_stale_indicator() -> None:
    try:
        app = PierForcesApp()
    except tk.TclError:
        pytest.skip("No display available for Tkinter smoke test.")

    try:
        app.withdraw()
        info = ETABSModelInfo(
            filename="Smoke.EDB",
            filepath=r"C:\Models\Smoke.EDB",
            units="kN, m, C",
            load_cases=["Dead", "Live", "EQx"],
            pier_labels=["P1"],
            stories=["Roof", "L2"],
            case_types={"Dead": "LinearStatic", "Live": "LinearStatic", "EQx": "LinearStaticMultiStep"},
        )
        raw = pd.DataFrame(
            [
                {"Story": "Roof", "Pier": "P1", "LoadCase": "Dead", "Location": "Top", "P": -100.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 100.0},
                {"Story": "Roof", "Pier": "P1", "LoadCase": "Dead", "Location": "Bottom", "P": -120.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 200.0},
                {"Story": "Roof", "Pier": "P1", "LoadCase": "Live", "Location": "Top", "P": -50.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 50.0},
                {"Story": "Roof", "Pier": "P1", "LoadCase": "Live", "Location": "Bottom", "P": -60.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 100.0},
                {"Story": "L2", "Pier": "P1", "LoadCase": "Dead", "Location": "Top", "P": -200.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": -100.0},
                {"Story": "L2", "Pier": "P1", "LoadCase": "Dead", "Location": "Bottom", "P": -220.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 400.0},
                {"Story": "L2", "Pier": "P1", "LoadCase": "Live", "Location": "Top", "P": -100.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": -50.0},
                {"Story": "L2", "Pier": "P1", "LoadCase": "Live", "Location": "Bottom", "P": -110.0, "V2": 0.0, "V3": 0.0, "T": 0.0, "M2": 0.0, "M3": 200.0},
            ]
        )

        # Default station filter is Bottom.
        assert app.station_var.get() == "Bottom"

        # Simulate connect + selecting Dead & Live + pull completion
        app._on_connect_done(info)
        app._case_vars["Dead"][0].set(True)
        app._case_vars["Dead"][1].set("Dead Load")
        app._case_vars["Live"][0].set(True)
        app._case_vars["Live"][1].set("Live Load")
        app._on_pull_done((info, raw), frozenset({"Dead", "Live"}))

        # 2 stories x 3 combos (Bottom only), with Cm3 computed from paired Top/Bottom!
        assert len(app._combos_df) == 6
        assert len(app._envelope_df) == 20  # 2 stories x 10 extreme rows
        assert set(app._combos_df["Location"]) == {"Bottom"}
        roof_14d = app._combos_df[
            (app._combos_df["Story"] == "Roof") & (app._combos_df["ComboName"] == "1.4D")
        ].iloc[0]
        assert roof_14d["Cm3"] == pytest.approx(0.80)  # single curvature: 140/280 = +0.5 -> 0.80
        l2_14d = app._combos_df[
            (app._combos_df["Story"] == "L2") & (app._combos_df["ComboName"] == "1.4D")
        ].iloc[0]
        assert l2_14d["Cm3"] == pytest.approx(0.50)  # double curvature: -140/560 = -0.25 -> 0.50

        # Test Level (Story) filter on Load Combinations & Envelope Summary tabs!
        app.combo_story_var.set("L2")
        filtered_combos = app._get_filtered_combos_df()
        assert len(filtered_combos) == 3
        assert set(filtered_combos["Story"]) == {"L2"}

        app.env_story_var.set("Roof")
        filtered_env = app._get_filtered_envelope_df()
        assert len(filtered_env) == 10
        assert set(filtered_env["Story"]) == {"Roof"}

        # Test Raw Imported Data tab (shows all 8 uncleaned Top+Bottom rows) + filters!
        assert len(app._get_filtered_raw_df()) == 8
        app.raw_story_var.set("Roof")
        app.raw_case_var.set("Dead")
        assert len(app._get_filtered_raw_df()) == 2
        app._reset_raw_filters()
        assert len(app._get_filtered_raw_df()) == 8

        # Test editable Combo Definitions tab: defaults -> custom edit -> recalculate -> reset!
        assert len(app.combo_defs_tree.get_children()) == 3  # 1.4D, 1.25D+1.5L, 0.9D+1.5L
        mutable = app._ensure_custom_combos_mutable()
        mutable.append(
            ComboDefinition(name="SLS_Check", factors={"Dead Load": 1.0, "Live Load": 1.0})
        )
        mutable[0].enabled = False  # disable 1.4D
        app._refresh_combo_defs_tree()
        app._on_recalculate()
        combo_names = set(app._combos_df["ComboName"])
        assert "1.4D" not in combo_names
        assert "SLS_Check" in combo_names

        app._on_reset_combo_defs()
        app._on_recalculate()
        assert "1.4D" in set(app._combos_df["ComboName"])
        assert "SLS_Check" not in set(app._combos_df["ComboName"])

        # Toggling an unpulled case triggers stale-cache warning
        app._case_vars["EQx"][0].set(True)
        assert "EQx" in app.stale_notice_var.get()

        # Unchecking EQx clears stale notice and Recalculate works from cache
        app._case_vars["EQx"][0].set(False)
        assert app.stale_notice_var.get() == ""
        app._on_recalculate()
        assert list(app._envelope_df.columns) == [
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
    finally:
        app.destroy()
