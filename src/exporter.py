"""Formatted 3-sheet Excel (.xlsx) workbook export and clipboard TSV formatting.

Pure Python/pandas/openpyxl (zero COM/ETABS dependencies).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

from src.config import AppConfig
from src.nbcc_combos import resolve_active_combo_definitions

SHEET_ENVELOPE = "Envelope Summary"
SHEET_COMBOS = "Load Combinations"
SHEET_METADATA = "Metadata"

_TABLE_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
_TABLE_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_SECTION_FILL = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="1F2937")
_SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="111827")
_BODY_FONT = Font(name="Calibri", size=10, color="111827")
_KEY_FONT = Font(name="Calibri", size=10, bold=True, color="1F2937")

_THIN_SIDE = Side(style="thin", color="BFBFBF")
_CELL_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE
)

_CM_COLUMNS: frozenset[str] = frozenset({"Cm3", "Cm2"})
_NUMERIC_COLUMNS: frozenset[str] = frozenset(
    {
        "P",
        "V2",
        "V3",
        "T",
        "M2",
        "M3",
        "Cm3",
        "Cm2",
    }
)


@dataclass
class ExportMetadata:
    """Audit metadata logged on the Metadata worksheet."""

    model_filename: str = "Unknown"
    model_filepath: str = ""
    units: str = "kN, m, C"
    exported_at: datetime = field(default_factory=datetime.now)
    raw_row_count: int = 0
    combo_row_count: int = 0
    envelope_row_count: int = 0


def excel_number_format(decimals: int) -> str:
    """Build an Excel numeric format string for `decimals` decimal places."""
    if decimals < 0:
        raise ValueError(f"decimals must be non-negative, got {decimals}.")
    if decimals == 0:
        return "#,##0"
    return f"#,##0.{'0' * decimals}"


def _auto_fit_columns(ws: Worksheet, min_width: int = 10, max_width: int = 50) -> None:
    for col_cells in ws.columns:
        col_idx = col_cells[0].column
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in col_cells:
            if cell.value is None:
                continue
            text = str(cell.value)
            # Account for formatted numeric display width if float/int.
            if isinstance(cell.value, float):
                text = f"{cell.value:,.2f}"
            max_len = max(max_len, len(text))
        ws.column_dimensions[col_letter].width = max(
            min_width, min(max_width, max_len + 4)
        )


def _write_data_sheet(
    ws: Worksheet,
    df: pd.DataFrame,
    decimals: int,
    table_name: str,
) -> None:
    """Write a DataFrame to an openpyxl worksheet as a formatted Excel Table."""
    num_fmt = excel_number_format(decimals)
    cm_fmt = excel_number_format(max(2, decimals))
    headers = list(df.columns)
    num_cols = len(headers)

    # Row 1: Styled header row.
    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = _TABLE_HEADER_FONT
        cell.fill = _TABLE_HEADER_FILL
        cell.border = _CELL_BORDER
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # Rows 2..N+1: Data rows (never overwriting row 1 or row 2).
    numeric_col_indices = {
        idx
        for idx, col_name in enumerate(headers, start=1)
        if col_name in _NUMERIC_COLUMNS
        or (col_name in df.columns and pd.api.types.is_numeric_dtype(df[col_name]))
    }
    cm_col_indices = {
        idx for idx, col_name in enumerate(headers, start=1) if col_name in _CM_COLUMNS
    }

    for row_offset, row_values in enumerate(df.itertuples(index=False, name=None), start=2):
        for col_idx, val in enumerate(row_values, start=1):
            if pd.isna(val):
                cell_val = None
            elif col_idx in numeric_col_indices:
                cell_val = float(val)
            else:
                cell_val = str(val)

            cell = ws.cell(row=row_offset, column=col_idx, value=cell_val)
            cell.font = _BODY_FONT
            cell.border = _CELL_BORDER
            if col_idx in numeric_col_indices:
                cell.number_format = cm_fmt if col_idx in cm_col_indices else num_fmt
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Freeze top header row.
    ws.freeze_panes = "A2"

    # Format populated sheets as native Excel Tables (blue header, banded rows, filter dropdowns).
    if num_cols > 0:
        last_col = get_column_letter(num_cols)
        last_row = len(df) + 1
        if len(df) > 0:
            table = Table(displayName=table_name, ref=f"A1:{last_col}{last_row}")
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            ws.add_table(table)
        else:
            ws.auto_filter.ref = f"A1:{last_col}1"

    _auto_fit_columns(ws)


def _write_section_banner(ws: Worksheet, row: int, title: str) -> int:
    cell_a = ws.cell(row=row, column=1, value=title)
    cell_b = ws.cell(row=row, column=2, value="")
    for cell in (cell_a, cell_b):
        cell.font = _SECTION_FONT
        cell.fill = _SECTION_FILL
        cell.border = _CELL_BORDER
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20
    return row + 1


def _write_kv_rows(ws: Worksheet, start_row: int, pairs: Iterable[tuple[str, object]]) -> int:
    row = start_row
    for key, val in pairs:
        c_key = ws.cell(row=row, column=1, value=key)
        c_val = ws.cell(row=row, column=2, value=val)
        c_key.font = _KEY_FONT
        c_val.font = _BODY_FONT
        c_key.border = _CELL_BORDER
        c_val.border = _CELL_BORDER
        c_key.alignment = Alignment(horizontal="left", vertical="center")
        c_val.alignment = Alignment(horizontal="left", vertical="center")
        row += 1
    return row


def _write_metadata_sheet(
    ws: Worksheet,
    config: AppConfig,
    metadata: ExportMetadata,
) -> None:
    """Write the structured audit trail onto the Metadata worksheet."""
    # Top header banner
    for col_idx, heading in enumerate(("Parameter", "Value"), start=1):
        cell = ws.cell(row=1, column=col_idx, value=heading)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.border = _CELL_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    row = 2
    row = _write_section_banner(ws, row, "Model Information")
    row = _write_kv_rows(
        ws,
        row,
        [
            ("Model Filename", metadata.model_filename),
            ("Model File Path", metadata.model_filepath or "N/A"),
            ("Export Timestamp", metadata.exported_at.strftime("%Y-%m-%d %H:%M:%S")),
            ("Active Output Units", metadata.units or config.units),
            ("Raw Rows Cached", metadata.raw_row_count),
            ("Load Combination Rows", metadata.combo_row_count),
            ("Envelope Summary Rows", metadata.envelope_row_count),
        ],
    )

    row += 1
    row = _write_section_banner(ws, row, "Load Case Mappings (ETABS -> NBCC 2020)")
    for col_idx, heading in enumerate(("ETABS Load Case", "NBCC Category"), start=1):
        cell = ws.cell(row=row, column=col_idx, value=heading)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.border = _CELL_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    row += 1

    if config.case_mappings:
        row = _write_kv_rows(ws, row, list(config.case_mappings.items()))
    else:
        row = _write_kv_rows(ws, row, [("(None)", "")])

    row += 1
    row = _write_section_banner(ws, row, "Cleaning & NBCC 2020 Settings")
    c = config.cleaning
    piers_str = "All" if c.selected_piers is None else ", ".join(c.selected_piers) or "(None)"
    stories_str = (
        "All" if c.selected_stories is None else ", ".join(c.selected_stories) or "(None)"
    )
    occupancy_str = (
        "Yes (Companion L = 1.0L in Cases 4 & 5)"
        if config.nbcc.storage_assembly_occupancy
        else "No (Standard: Companion L = 0.5L in Cases 4 & 5)"
    )

    combo_mode = (
        "Custom / Edited"
        if config.nbcc.custom_combinations is not None
        else "NBCC 2020 Table 4.1.3.2.-A (Default)"
    )

    row = _write_kv_rows(
        ws,
        row,
        [
            ("Pier Filter", piers_str),
            ("Story Filter", stories_str),
            ("Station Filter", c.station_filter),
            ("Decimal Rounding", c.round_decimals),
            ("Drop Near-Zero Rows", "Yes" if c.drop_near_zero else "No"),
            ("Near-Zero Threshold", c.near_zero_threshold),
            ("Remove Duplicate Rows", "Yes" if c.remove_duplicates else "No"),
            ("Aggregate Gravity Sub-Cases (Dead, Live, Snow)", "Yes" if c.aggregate_subcases else "No"),
            ("Zero Lateral Axial P (Wind, Seismic)", "Yes" if c.zero_lateral_axial else "No"),
            (
                "Zero Gravity Shear & Moment (Dead, Live, Snow)",
                "Yes" if c.zero_gravity_shear_moment else "No",
            ),
            ("Storage / Assembly Occupancy", occupancy_str),
            ("Load Combination Mode", combo_mode),
        ],
    )

    row += 1
    row = _write_section_banner(ws, row, "Active Load Combination Definitions")
    for col_idx, heading in enumerate(("Combo Name", "Category Factors"), start=1):
        cell = ws.cell(row=row, column=col_idx, value=heading)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.border = _CELL_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    row += 1

    mapped_cats = set(config.case_mappings.values())
    active_defs = resolve_active_combo_definitions(mapped_cats, config.nbcc)
    if active_defs:
        def_rows = [
            (
                combo.name,
                ", ".join(f"{cat}={fac:g}" for cat, fac in combo.factors.items()),
            )
            for combo in active_defs
        ]
        _write_kv_rows(ws, row, def_rows)
    else:
        _write_kv_rows(ws, row, [("(None)", "")])

    _auto_fit_columns(ws, min_width=18, max_width=70)


def export_to_excel(
    envelope_df: pd.DataFrame,
    combos_df: pd.DataFrame,
    config: AppConfig,
    metadata: ExportMetadata,
    output_path: str | Path,
) -> Path:
    """Export Envelope Summary, Load Combinations, and Metadata to one .xlsx workbook."""
    config.validate()
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()

    ws_env = wb.active
    ws_env.title = SHEET_ENVELOPE
    _write_data_sheet(
        ws_env,
        envelope_df,
        config.cleaning.round_decimals,
        table_name="Envelope_Summary",
    )

    ws_combos = wb.create_sheet(title=SHEET_COMBOS)
    _write_data_sheet(
        ws_combos,
        combos_df,
        config.cleaning.round_decimals,
        table_name="Load_Combinations",
    )

    ws_meta = wb.create_sheet(title=SHEET_METADATA)
    _write_metadata_sheet(ws_meta, config, metadata)

    wb.save(target)
    return target


def dataframe_to_tsv(df: pd.DataFrame, decimals: int = 2) -> str:
    """Format a DataFrame as tab-separated text suitable for pasting into Excel."""
    if decimals < 0:
        raise ValueError(f"decimals must be non-negative, got {decimals}.")
    if df.empty and len(df.columns) == 0:
        return ""

    headers = "\t".join(str(col) for col in df.columns)
    if df.empty:
        return headers

    fmt = f"{{:.{decimals}f}}"
    cm_fmt = f"{{:.{max(2, decimals)}f}}"
    numeric_cols = {
        idx
        for idx, col in enumerate(df.columns)
        if col in _NUMERIC_COLUMNS or pd.api.types.is_numeric_dtype(df[col])
    }
    cm_cols = {idx for idx, col in enumerate(df.columns) if col in _CM_COLUMNS}

    lines = [headers]
    for row_vals in df.itertuples(index=False, name=None):
        formatted = []
        for idx, val in enumerate(row_vals):
            if pd.isna(val):
                formatted.append("")
            elif idx in numeric_cols:
                num = float(val)
                if num == 0.0:
                    num = 0.0
                formatted.append((cm_fmt if idx in cm_cols else fmt).format(num))
            else:
                formatted.append(str(val))
        lines.append("\t".join(formatted))
    return "\n".join(lines)
