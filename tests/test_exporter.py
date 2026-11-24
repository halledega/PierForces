"""Unit tests for src.exporter."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

from src.config import AppConfig, CleaningOptions, NBCCOptions
from src.exporter import (
    SHEET_COMBOS,
    SHEET_ENVELOPE,
    SHEET_METADATA,
    ExportMetadata,
    dataframe_to_tsv,
    excel_number_format,
    export_to_excel,
)
from src.nbcc_combos import run_nbcc_analysis


def test_export_to_excel_three_sheets_formatting_and_no_row_overwrite(
    tmp_path: Path,
) -> None:
    cleaned = pd.DataFrame(
        [
            {"Story": "L2", "Pier": "P1", "LoadCase": "D", "Category": "D", "Location": "Bottom", "P": -500.125, "V2": 5.0, "V3": 1.0, "T": 0.5, "M2": 2.0, "M3": 10.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "L", "Category": "L", "Location": "Bottom", "P": -200.0, "V2": 2.0, "V3": 0.5, "T": 0.2, "M2": 1.0, "M3": 4.0},
            {"Story": "L2", "Pier": "P1", "LoadCase": "EQx", "Category": "E", "Location": "Bottom", "P": 80.0, "V2": 300.0, "V3": 10.0, "T": 20.0, "M2": 40.0, "M3": 1500.0},
        ]
    )
    cfg = AppConfig(
        units="kN, m, C",
        case_mappings={"DL": "D", "LIVE": "L", "EQx": "E"},
        cleaning=CleaningOptions(
            selected_piers=["P1"],
            selected_stories=["L2"],
            station_filter="Bottom",
            round_decimals=2,
            drop_near_zero=True,
            near_zero_threshold=1e-4,
            aggregate_subcases=True,
        ),
        nbcc=NBCCOptions(storage_assembly_occupancy=True),
    )

    combos_df, env_df = run_nbcc_analysis(cleaned, cfg.cleaning, cfg.nbcc)
    meta = ExportMetadata(
        model_filename="SampleModel_v1.2.EDB",
        model_filepath=r"C:\Models\SampleModel_v1.2.EDB",
        units=cfg.units,
        exported_at=datetime(2026, 11, 23, 12, 30, 0),
        raw_row_count=12,
        combo_row_count=len(combos_df),
        envelope_row_count=len(env_df),
    )

    out_path = tmp_path / "pier_forces.xlsx"
    export_to_excel(env_df, combos_df, cfg, meta, out_path)

    wb = openpyxl.load_workbook(out_path)
    assert wb.sheetnames == [SHEET_ENVELOPE, SHEET_COMBOS, SHEET_METADATA]

    # Envelope Summary sheet verification (Excel Table, freeze panes, exact 13 columns, P Min intact!).
    ws_env = wb[SHEET_ENVELOPE]
    assert ws_env.freeze_panes == "A2"
    assert "Envelope_Summary" in ws_env.tables
    assert ws_env.tables["Envelope_Summary"].ref == f"A1:M{len(env_df) + 1}"
    assert [ws_env.cell(1, c).value for c in range(1, 14)] == [
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
    assert ws_env.cell(1, 1).font.bold is True
    assert ws_env.cell(2, 1).value == "P1"  # First data row intact!
    assert ws_env.cell(2, 2).value == "L2"
    assert ws_env.cell(2, 4).value == "P Min"  # P Min never overwritten!
    assert ws_env.cell(2, 6).number_format == "#,##0.00"
    assert ws_env.cell(2, 10).number_format == "#,##0.00"  # Cm3
    assert ws_env.cell(2, 13).number_format == "#,##0.00"  # Cm2
    assert ws_env.max_row == len(env_df) + 1

    # Load Combinations sheet verification (Excel Table, freeze panes, 12 columns, no overwrite).
    ws_combos = wb[SHEET_COMBOS]
    assert ws_combos.freeze_panes == "A2"
    assert "Load_Combinations" in ws_combos.tables
    assert ws_combos.tables["Load_Combinations"].ref == f"A1:L{len(combos_df) + 1}"
    assert ws_combos.cell(1, 4).value == "ComboName"
    assert ws_combos.cell(1, 9).value == "Cm3"
    assert ws_combos.cell(1, 12).value == "Cm2"
    assert ws_combos.cell(2, 4).value == "1.4D"  # First combo row intact!
    assert ws_combos.cell(2, 5).number_format == "#,##0.00"
    assert ws_combos.max_row == len(combos_df) + 1

    # Metadata sheet audit trail verification.
    ws_meta = wb[SHEET_METADATA]
    meta_values = {
        ws_meta.cell(r, 1).value: ws_meta.cell(r, 2).value
        for r in range(1, ws_meta.max_row + 1)
        if ws_meta.cell(r, 1).value is not None
    }
    assert meta_values["Model Filename"] == "SampleModel_v1.2.EDB"
    assert meta_values["Active Output Units"] == "kN, m, C"
    assert meta_values["DL"] == "Dead Load"
    assert meta_values["EQx"] == "Seismic Load"
    assert meta_values["Station Filter"] == "Bottom"
    assert meta_values["Decimal Rounding"] == 2
    assert "1.0L" in str(meta_values["Storage / Assembly Occupancy"])
    assert meta_values["Load Combination Mode"] == "NBCC 2020 Table 4.1.3.2.-A (Default)"
    assert meta_values["1.4D"] == "Dead Load=1.4"


def test_excel_number_format_and_tsv() -> None:
    assert excel_number_format(0) == "#,##0"
    assert excel_number_format(2) == "#,##0.00"
    assert excel_number_format(3) == "#,##0.000"

    df = pd.DataFrame(
        [
            {"Story": "L2", "Pier": "P1", "P": -100.5, "M3": 25.125},
        ]
    )
    tsv = dataframe_to_tsv(df, decimals=2)
    assert tsv == "Story\tPier\tP\tM3\nL2\tP1\t-100.50\t25.12"
