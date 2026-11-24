"""Tkinter/ttk desktop interface for ETABS Pier Forces (NBCC 2020).

Manages the 4-tab workflow, in-memory raw pier-force caching, background
threading for ETABS COM operations, JSON preset save/load, Excel export, and
clipboard copy. All ETABS COM calls are delegated strictly to `src.etabs_client`.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

import pandas as pd

from src.cleaner import clean_pier_forces
from src.config import (
    AppConfig,
    CleaningOptions,
    ComboDefinition,
    DEFAULT_UNITS,
    NBCC_CATEGORIES,
    NBCCOptions,
    STATION_FILTERS,
    UNIT_SYSTEMS,
    load_config,
    reconcile_config_with_model,
    save_config,
)
from src.etabs_client import ETABSClient, ETABSError, ETABSModelInfo
from src.exporter import ExportMetadata, dataframe_to_tsv, export_to_excel
from src.nbcc_combos import (
    default_combo_definitions,
    format_combo_name,
    run_nbcc_analysis,
)


def _suggest_nbcc_category(case_name: str) -> str:
    """Best-effort heuristic suggestion for an ETABS load case name."""
    upper = case_name.upper()
    if any(tok in upper for tok in ("EQ", "SEIS", "SPEC", "RSX", "RSY", "RS_", "EX", "EY")):
        return "Seismic Load"
    if any(tok in upper for tok in ("WIND", "WX", "WY", "WL")):
        return "Wind Load"
    if any(tok in upper for tok in ("SNOW", "SL")):
        return "Snow Load"
    if any(tok in upper for tok in ("LIVE", "LL", "STORAGE", "OCC", "ROOF_L")):
        return "Live Load"
    return "Dead Load"


class _ScrollableFrame(ttk.Frame):
    """Vertically scrollable ttk container for dynamic load-case mapping rows."""

    def __init__(self, parent: tk.Widget, **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self._window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scroll.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _on_inner_configure(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self._window_id, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_mousewheel(self, _event: tk.Event) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event: tk.Event) -> None:
        self.canvas.unbind_all("<MouseWheel>")


class _ComboEditorDialog(tk.Toplevel):
    """Modal dialog for creating or editing a category-level ComboDefinition."""

    def __init__(
        self,
        parent: tk.Widget,
        title: str,
        initial: ComboDefinition | None = None,
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.result: ComboDefinition | None = None
        self._case_id = initial.case_id if initial is not None else "Custom"

        self._name_var = tk.StringVar(value=initial.name if initial else "")
        self._enabled_var = tk.BooleanVar(value=initial.enabled if initial else True)
        self._factor_vars: dict[str, tk.StringVar] = {}
        for cat in NBCC_CATEGORIES:
            val = initial.factors.get(cat, 0.0) if initial else 0.0
            self._factor_vars[cat] = tk.StringVar(value="" if val == 0.0 else f"{val:g}")

        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Combination Name:").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        name_row = ttk.Frame(body)
        name_row.grid(row=0, column=1, sticky="ew", pady=(0, 6), padx=(8, 0))
        ttk.Entry(name_row, textvariable=self._name_var, width=28).pack(side="left")
        ttk.Button(
            name_row, text="Auto-Name", command=self._auto_name
        ).pack(side="left", padx=(6, 0))

        ttk.Checkbutton(
            body, text="Enabled (include in calculations)", variable=self._enabled_var
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        factors_box = ttk.LabelFrame(
            body, text="Category Load Factors (leave blank or 0 to omit)", padding=10
        )
        factors_box.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 12))

        labels = {
            "Dead Load": "Dead Load (D)",
            "Live Load": "Live Load (L)",
            "Snow Load": "Snow Load (S)",
            "Wind Load": "Wind Load (W, +/-)",
            "Seismic Load": "Seismic Load (E, +/-)",
        }
        for idx, cat in enumerate(NBCC_CATEGORIES):
            ttk.Label(factors_box, text=labels[cat]).grid(
                row=idx, column=0, sticky="w", pady=3, padx=(0, 12)
            )
            ttk.Entry(
                factors_box, textvariable=self._factor_vars[cat], width=14
            ).grid(row=idx, column=1, sticky="w", pady=3)

        btns = ttk.Frame(body)
        btns.grid(row=3, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(
            side="right", padx=(6, 0)
        )
        ttk.Button(btns, text="OK", style="Accent.TButton", command=self._on_ok).pack(
            side="right"
        )

        self.bind("<Return>", lambda _e: self._on_ok())
        self.bind("<Escape>", lambda _e: self.destroy())
        self.wait_visibility()
        self.focus_set()

    def _parse_factors(self) -> dict[str, float]:
        factors: dict[str, float] = {}
        for cat in NBCC_CATEGORIES:
            raw = self._factor_vars[cat].get().strip()
            if not raw:
                continue
            try:
                val = float(raw)
            except ValueError as exc:
                raise ValueError(
                    f"Factor for {cat} must be a valid number, got {raw!r}."
                ) from exc
            if val != 0.0:
                factors[cat] = val
        return factors

    def _auto_name(self) -> None:
        try:
            factors = self._parse_factors()
        except ValueError as exc:
            messagebox.showerror("Invalid Factor", str(exc), parent=self)
            return
        if not factors:
            messagebox.showwarning(
                "No Factors", "Enter at least one non-zero factor first.", parent=self
            )
            return
        self._name_var.set(format_combo_name(list(factors.items())))

    def _on_ok(self) -> None:
        try:
            factors = self._parse_factors()
            name = self._name_var.get().strip()
            if not name and factors:
                name = format_combo_name(list(factors.items()))
            combo = ComboDefinition(
                name=name,
                factors=factors,
                enabled=bool(self._enabled_var.get()),
                case_id=self._case_id,
            )
            combo.validate()
        except ValueError as exc:
            messagebox.showerror("Invalid Combination", str(exc), parent=self)
            return
        self.result = combo
        self.destroy()


class PierForcesApp(tk.Tk):
    """Main desktop application window for ETABS Pier Forces (NBCC 2020)."""

    def __init__(self, client: ETABSClient | None = None) -> None:
        super().__init__()
        self.title("ETABS Pier Forces — NBCC 2020 ULS")
        self.geometry("1260x780")
        self.minsize(1020, 660)

        self._client = client or ETABSClient()
        self._busy = False

        # Connected ETABS state & in-memory cache
        self._model_info: ETABSModelInfo | None = None
        self._pending_config: AppConfig | None = None
        self._raw_df: pd.DataFrame | None = None
        self._pulled_cases: frozenset[str] = frozenset()
        self._pulled_units: str | None = None
        self._pulled_model_name: str = ""
        self._combos_df: pd.DataFrame = pd.DataFrame()
        self._envelope_df: pd.DataFrame = pd.DataFrame()
        self._custom_combos: list[ComboDefinition] | None = None

        # Dynamic load-case mapping row vars: case_name -> (selected_var, category_var)
        self._case_vars: dict[str, tuple[tk.BooleanVar, tk.StringVar]] = {}

        # Form state variables
        self.units_var = tk.StringVar(value=DEFAULT_UNITS)
        self.model_label_var = tk.StringVar(value="Not connected to ETABS")
        self.model_counts_var = tk.StringVar(value="Load Cases: — | Piers: — | Stories: —")
        self.stale_notice_var = tk.StringVar(value="")

        self.station_var = tk.StringVar(value="Bottom")
        self.decimals_var = tk.IntVar(value=2)
        self.drop_near_zero_var = tk.BooleanVar(value=False)
        self.near_zero_threshold_var = tk.StringVar(value="1e-4")
        self.remove_duplicates_var = tk.BooleanVar(value=False)
        self.aggregate_subcases_var = tk.BooleanVar(value=True)
        self.zero_lateral_axial_var = tk.BooleanVar(value=False)
        self.zero_gravity_shear_moment_var = tk.BooleanVar(value=False)
        self.storage_occupancy_var = tk.BooleanVar(value=False)

        self.combo_defs_mode_var = tk.StringVar(
            value="Mode: NBCC 2020 Table 4.1.3.2.-A Defaults (0 combos)"
        )
        self.raw_story_var = tk.StringVar(value="All Stories")
        self.raw_pier_var = tk.StringVar(value="All Piers")
        self.raw_case_var = tk.StringVar(value="All Cases")
        self.raw_loc_var = tk.StringVar(value="All")
        self.raw_count_var = tk.StringVar(value="0 raw rows imported")

        self.combo_story_var = tk.StringVar(value="All Stories")
        self.combo_pier_var = tk.StringVar(value="All Piers")
        self.combo_search_var = tk.StringVar(value="")
        self.combo_count_var = tk.StringVar(value="0 combinations")
        self.env_story_var = tk.StringVar(value="All Stories")
        self.env_pier_var = tk.StringVar(value="All Piers")
        self.envelope_count_var = tk.StringVar(value="0 envelope rows")
        self.status_var = tk.StringVar(value="Ready. Click 'Connect to ETABS' to begin.")

        self._configure_style()
        self._build_toolbar()
        self._build_notebook()
        self._build_status_bar()

        # Watch units & occupancy changes
        self.units_var.trace_add("write", lambda *_: self._update_stale_indicator())
        self.storage_occupancy_var.trace_add(
            "write", lambda *_: self._refresh_combo_defs_tree()
        )
        self._refresh_combo_defs_tree()

    # ------------------------------------------------------------------
    # UI Construction
    # ------------------------------------------------------------------
    def _configure_style(self) -> None:
        style = ttk.Style(self)
        for theme in ("vista", "xpnative", "clam", "default"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure("Toolbar.TButton", padding=(10, 5))
        style.configure("Accent.TButton", padding=(12, 6), font=("Segoe UI", 9, "bold"))
        style.configure("Heading.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Warn.TLabel", foreground="#B45309", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=23, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self, padding=(10, 8, 10, 4))
        bar.pack(fill="x")

        self.btn_connect = ttk.Button(
            bar, text="Connect to ETABS", style="Accent.TButton", command=self._on_connect
        )
        self.btn_connect.pack(side="left", padx=(0, 6))

        self.btn_pull = ttk.Button(
            bar, text="Pull from ETABS", style="Accent.TButton", command=self._on_pull
        )
        self.btn_pull.pack(side="left", padx=(0, 6))

        self.btn_recalc_top = ttk.Button(
            bar, text="Recalculate", style="Toolbar.TButton", command=self._on_recalculate
        )
        self.btn_recalc_top.pack(side="left", padx=(0, 12))

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=4, pady=2)

        self.btn_save_cfg = ttk.Button(
            bar, text="Save Config", style="Toolbar.TButton", command=self._on_save_config
        )
        self.btn_save_cfg.pack(side="left", padx=(8, 4))

        self.btn_load_cfg = ttk.Button(
            bar, text="Load Config", style="Toolbar.TButton", command=self._on_load_config
        )
        self.btn_load_cfg.pack(side="left", padx=(0, 12))

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=4, pady=2)

        self.btn_export = ttk.Button(
            bar,
            text="Export to Excel (.xlsx)",
            style="Toolbar.TButton",
            command=self._on_export_excel,
        )
        self.btn_export.pack(side="left", padx=(8, 4))

        self.btn_copy = ttk.Button(
            bar,
            text="Copy to Clipboard",
            style="Toolbar.TButton",
            command=self._on_copy_clipboard,
        )
        self.btn_copy.pack(side="left", padx=(0, 6))

        # Active model label on right side of toolbar
        ttk.Label(bar, textvariable=self.model_label_var, style="Heading.TLabel").pack(
            side="right", padx=(10, 4)
        )

    def _build_notebook(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(4, 6))

        self.tab_setup = ttk.Frame(self.notebook, padding=10)
        self.tab_options = ttk.Frame(self.notebook, padding=10)
        self.tab_combo_defs = ttk.Frame(self.notebook, padding=10)
        self.tab_raw = ttk.Frame(self.notebook, padding=10)
        self.tab_combos = ttk.Frame(self.notebook, padding=10)
        self.tab_envelope = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_setup, text="1. Setup & Case Mapping")
        self.notebook.add(self.tab_options, text="2. Cleaning & NBCC Options")
        self.notebook.add(self.tab_combo_defs, text="3. Combo Definitions")
        self.notebook.add(self.tab_raw, text="4. Raw Imported Data")
        self.notebook.add(self.tab_combos, text="5. Load Combinations")
        self.notebook.add(self.tab_envelope, text="6. Envelope Summary")

        self._build_setup_tab()
        self._build_options_tab()
        self._build_combo_defs_tab()
        self._build_raw_tab()
        self._build_combos_tab()
        self._build_envelope_tab()

    def _build_setup_tab(self) -> None:
        top = ttk.LabelFrame(self.tab_setup, text="ETABS Model & Output Units", padding=10)
        top.pack(fill="x", pady=(0, 10))

        row1 = ttk.Frame(top)
        row1.pack(fill="x")

        ttk.Label(row1, text="Output Units:").pack(side="left")
        units_cb = ttk.Combobox(
            row1,
            textvariable=self.units_var,
            values=list(UNIT_SYSTEMS),
            state="readonly",
            width=14,
        )
        units_cb.pack(side="left", padx=(6, 18))

        ttk.Label(row1, textvariable=self.model_counts_var).pack(side="left")

        ttk.Label(
            top,
            textvariable=self.stale_notice_var,
            style="Warn.TLabel",
        ).pack(anchor="w", pady=(6, 0))

        # Load case mapping frame
        map_frame = ttk.LabelFrame(
            self.tab_setup,
            text="Select Load Cases & Map to NBCC 2020 Load Types (Dead / Live / Snow / Wind / Seismic)",
            padding=10,
        )
        map_frame.pack(fill="both", expand=True)

        actions = ttk.Frame(map_frame)
        actions.pack(fill="x", pady=(0, 8))

        ttk.Button(actions, text="Select All", command=self._select_all_cases).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(actions, text="Clear All", command=self._clear_all_cases).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(
            actions,
            text="Auto-Suggest Load Types",
            command=self._auto_suggest_categories,
        ).pack(side="left", padx=(0, 12))
        ttk.Label(
            actions,
            text="Tip: Multiple gravity cases (e.g. DL + SDL -> Dead Load) can share a type. "
            "Directional Wind/Seismic cases stay separate with +/- reversals.",
        ).pack(side="left")

        header = ttk.Frame(map_frame)
        header.pack(fill="x", padx=(4, 20), pady=(0, 4))
        ttk.Label(header, text="Include", width=10, style="Heading.TLabel").pack(side="left")
        ttk.Label(header, text="ETABS Load Case", width=34, style="Heading.TLabel").pack(
            side="left"
        )
        ttk.Label(header, text="ETABS Case Type", width=24, style="Heading.TLabel").pack(
            side="left"
        )
        ttk.Label(header, text="NBCC Load Type", width=18, style="Heading.TLabel").pack(
            side="left"
        )

        self.case_scroll = _ScrollableFrame(map_frame)
        self.case_scroll.pack(fill="both", expand=True)

        self._render_empty_case_placeholder()

    def _render_empty_case_placeholder(self) -> None:
        for widget in self.case_scroll.inner.winfo_children():
            widget.destroy()
        ttk.Label(
            self.case_scroll.inner,
            text="Connect to an open ETABS model to list available load cases.",
            padding=12,
        ).pack(anchor="w")

    def _build_options_tab(self) -> None:
        container = ttk.Frame(self.tab_options)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=1)
        container.rowconfigure(0, weight=1)

        # Column 0: Pier Label Multi-Select
        pier_box = ttk.LabelFrame(container, text="Pier Label Filter", padding=8)
        pier_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        pier_btns = ttk.Frame(pier_box)
        pier_btns.pack(fill="x", pady=(0, 6))
        ttk.Button(
            pier_btns, text="Select All", command=lambda: self._select_all_listbox(self.pier_listbox)
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            pier_btns, text="Clear All", command=lambda: self._clear_listbox(self.pier_listbox)
        ).pack(side="left")

        pier_list_frame = ttk.Frame(pier_box)
        pier_list_frame.pack(fill="both", expand=True)
        self.pier_listbox = tk.Listbox(
            pier_list_frame, selectmode="extended", exportselection=False
        )
        pier_scroll = ttk.Scrollbar(
            pier_list_frame, orient="vertical", command=self.pier_listbox.yview
        )
        self.pier_listbox.configure(yscrollcommand=pier_scroll.set)
        self.pier_listbox.pack(side="left", fill="both", expand=True)
        pier_scroll.pack(side="right", fill="y")

        # Column 1: Story Multi-Select
        story_box = ttk.LabelFrame(container, text="Story Filter (Model Order)", padding=8)
        story_box.grid(row=0, column=1, sticky="nsew", padx=(0, 8))

        story_btns = ttk.Frame(story_box)
        story_btns.pack(fill="x", pady=(0, 6))
        ttk.Button(
            story_btns,
            text="Select All",
            command=lambda: self._select_all_listbox(self.story_listbox),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            story_btns,
            text="Clear All",
            command=lambda: self._clear_listbox(self.story_listbox),
        ).pack(side="left")

        story_list_frame = ttk.Frame(story_box)
        story_list_frame.pack(fill="both", expand=True)
        self.story_listbox = tk.Listbox(
            story_list_frame, selectmode="extended", exportselection=False
        )
        story_scroll = ttk.Scrollbar(
            story_list_frame, orient="vertical", command=self.story_listbox.yview
        )
        self.story_listbox.configure(yscrollcommand=story_scroll.set)
        self.story_listbox.pack(side="left", fill="both", expand=True)
        story_scroll.pack(side="right", fill="y")

        # Column 2: Station, Cleaning, NBCC, Recalculate
        right = ttk.Frame(container)
        right.grid(row=0, column=2, sticky="nsew")

        station_frame = ttk.LabelFrame(right, text="Station Filter", padding=8)
        station_frame.pack(fill="x", pady=(0, 8))
        for st in STATION_FILTERS:
            ttk.Radiobutton(
                station_frame, text=st, value=st, variable=self.station_var
            ).pack(side="left", padx=(0, 14))

        clean_frame = ttk.LabelFrame(right, text="Data Cleaning & Precision", padding=8)
        clean_frame.pack(fill="x", pady=(0, 8))

        dec_row = ttk.Frame(clean_frame)
        dec_row.pack(fill="x", pady=(0, 6))
        ttk.Label(dec_row, text="Decimal places for outputs:").pack(side="left")
        ttk.Spinbox(
            dec_row, from_=0, to=6, width=6, textvariable=self.decimals_var
        ).pack(side="left", padx=(8, 0))

        nz_row = ttk.Frame(clean_frame)
        nz_row.pack(fill="x", pady=(0, 4))
        ttk.Checkbutton(
            nz_row,
            text="Drop near-zero rows (threshold):",
            variable=self.drop_near_zero_var,
        ).pack(side="left")
        ttk.Entry(nz_row, textvariable=self.near_zero_threshold_var, width=10).pack(
            side="left", padx=(6, 0)
        )

        ttk.Checkbutton(
            clean_frame,
            text="Remove exact duplicate rows",
            variable=self.remove_duplicates_var,
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            clean_frame,
            text="Aggregate gravity sub-cases (Dead, Live, Snow) into single totals",
            variable=self.aggregate_subcases_var,
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            clean_frame,
            text="Zero axial force (P) for lateral cases (Wind, Seismic)",
            variable=self.zero_lateral_axial_var,
        ).pack(anchor="w", pady=2)

        ttk.Checkbutton(
            clean_frame,
            text="Zero shear & moments (V2, V3, T, M2, M3) for gravity cases (Dead, Live, Snow)",
            variable=self.zero_gravity_shear_moment_var,
        ).pack(anchor="w", pady=2)

        nbcc_frame = ttk.LabelFrame(right, text="NBCC 2020 Table 4.1.3.2.-A Options", padding=8)
        nbcc_frame.pack(fill="x", pady=(0, 10))
        ttk.Checkbutton(
            nbcc_frame,
            text="Storage / Assembly Occupancy (companion L: 0.5L -> 1.0L in Cases 4 & 5)",
            variable=self.storage_occupancy_var,
        ).pack(anchor="w", pady=2)

        self.btn_recalc_tab = ttk.Button(
            right,
            text="Recalculate from Cached Data",
            style="Accent.TButton",
            command=self._on_recalculate,
        )
        self.btn_recalc_tab.pack(fill="x", pady=(4, 4))

        ttk.Label(
            right,
            textvariable=self.stale_notice_var,
            style="Warn.TLabel",
            wraplength=360,
        ).pack(anchor="w")

    def _build_combo_defs_tab(self) -> None:
        header = ttk.Frame(self.tab_combo_defs)
        header.pack(fill="x", pady=(0, 6))
        ttk.Label(
            header, textvariable=self.combo_defs_mode_var, style="Heading.TLabel"
        ).pack(side="left")

        btns = ttk.Frame(self.tab_combo_defs)
        btns.pack(fill="x", pady=(0, 8))
        ttk.Button(btns, text="Add Combo", command=self._on_add_combo_def).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(btns, text="Edit Selected", command=self._on_edit_combo_def).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(
            btns, text="Toggle Enable", command=self._on_toggle_combo_def
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            btns, text="Delete Selected", command=self._on_delete_combo_def
        ).pack(side="left", padx=(0, 12))
        ttk.Button(
            btns,
            text="Reset to NBCC 2020 Defaults",
            command=self._on_reset_combo_defs,
        ).pack(side="left", padx=(0, 12))
        ttk.Button(
            btns,
            text="Apply & Recalculate",
            style="Accent.TButton",
            command=self._on_recalculate,
        ).pack(side="right")

        ttk.Label(
            self.tab_combo_defs,
            text="Double-click any row to edit its factors/name, or press Space to enable/disable. "
            "Directional Wind/Seismic cases and multi-step seismic cases expand automatically.",
        ).pack(anchor="w", pady=(0, 6))

        self.combo_defs_tree = self._create_treeview(self.tab_combo_defs)
        cols = ["Enabled", "Case", "Combo Name", *NBCC_CATEGORIES]
        self.combo_defs_tree["columns"] = cols
        widths = {
            "Enabled": 72,
            "Case": 130,
            "Combo Name": 240,
            **{cat: 96 for cat in NBCC_CATEGORIES},
        }
        for col in cols:
            anchor = "e" if col in NBCC_CATEGORIES else ("center" if col == "Enabled" else "w")
            self.combo_defs_tree.heading(col, text=col, anchor="center")
            self.combo_defs_tree.column(
                col, width=widths[col], minwidth=64, anchor=anchor, stretch=(col == "Combo Name")
            )

        self.combo_defs_tree.bind("<Double-1>", lambda _e: self._on_edit_combo_def())
        self.combo_defs_tree.bind("<space>", lambda _e: self._on_toggle_combo_def())

    def _build_raw_tab(self) -> None:
        top = ttk.Frame(self.tab_raw)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, textvariable=self.raw_count_var, style="Heading.TLabel").pack(
            side="left"
        )

        ttk.Label(top, text="Level (Story):").pack(side="left", padx=(16, 4))
        self.raw_story_cb = ttk.Combobox(
            top,
            textvariable=self.raw_story_var,
            values=["All Stories"],
            state="readonly",
            width=13,
        )
        self.raw_story_cb.pack(side="left")

        ttk.Label(top, text="Pier:").pack(side="left", padx=(10, 4))
        self.raw_pier_cb = ttk.Combobox(
            top,
            textvariable=self.raw_pier_var,
            values=["All Piers"],
            state="readonly",
            width=11,
        )
        self.raw_pier_cb.pack(side="left")

        ttk.Label(top, text="Load Case:").pack(side="left", padx=(10, 4))
        self.raw_case_cb = ttk.Combobox(
            top,
            textvariable=self.raw_case_var,
            values=["All Cases"],
            state="readonly",
            width=15,
        )
        self.raw_case_cb.pack(side="left")

        ttk.Label(top, text="Location:").pack(side="left", padx=(10, 4))
        self.raw_loc_cb = ttk.Combobox(
            top,
            textvariable=self.raw_loc_var,
            values=["All", "Top", "Bottom"],
            state="readonly",
            width=9,
        )
        self.raw_loc_cb.pack(side="left")

        ttk.Button(top, text="Reset", command=self._reset_raw_filters).pack(
            side="left", padx=(8, 0)
        )

        for var in (
            self.raw_story_var,
            self.raw_pier_var,
            self.raw_case_var,
            self.raw_loc_var,
        ):
            var.trace_add("write", lambda *_: self._refresh_raw_tree())

        self.raw_tree = self._create_treeview(self.tab_raw)

    def _build_combos_tab(self) -> None:
        top = ttk.Frame(self.tab_combos)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, textvariable=self.combo_count_var, style="Heading.TLabel").pack(
            side="left"
        )

        ttk.Label(top, text="Level (Story):").pack(side="left", padx=(18, 4))
        self.combo_story_cb = ttk.Combobox(
            top,
            textvariable=self.combo_story_var,
            values=["All Stories"],
            state="readonly",
            width=14,
        )
        self.combo_story_cb.pack(side="left")

        ttk.Label(top, text="Pier:").pack(side="left", padx=(12, 4))
        self.combo_pier_cb = ttk.Combobox(
            top,
            textvariable=self.combo_pier_var,
            values=["All Piers"],
            state="readonly",
            width=12,
        )
        self.combo_pier_cb.pack(side="left")

        ttk.Label(top, text="Search:").pack(side="left", padx=(12, 4))
        search_entry = ttk.Entry(top, textvariable=self.combo_search_var, width=22)
        search_entry.pack(side="left")
        ttk.Button(
            top, text="Reset", command=self._reset_combo_filters
        ).pack(side="left", padx=(6, 0))

        self.combo_story_var.trace_add("write", lambda *_: self._refresh_combos_tree())
        self.combo_pier_var.trace_add("write", lambda *_: self._refresh_combos_tree())
        self.combo_search_var.trace_add("write", lambda *_: self._refresh_combos_tree())

        self.combos_tree = self._create_treeview(self.tab_combos)

    def _build_envelope_tab(self) -> None:
        top = ttk.Frame(self.tab_envelope)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, textvariable=self.envelope_count_var, style="Heading.TLabel").pack(
            side="left"
        )

        ttk.Label(top, text="Level (Story):").pack(side="left", padx=(18, 4))
        self.env_story_cb = ttk.Combobox(
            top,
            textvariable=self.env_story_var,
            values=["All Stories"],
            state="readonly",
            width=14,
        )
        self.env_story_cb.pack(side="left")

        ttk.Label(top, text="Pier:").pack(side="left", padx=(12, 4))
        self.env_pier_cb = ttk.Combobox(
            top,
            textvariable=self.env_pier_var,
            values=["All Piers"],
            state="readonly",
            width=12,
        )
        self.env_pier_cb.pack(side="left")

        ttk.Button(
            top, text="Reset", command=self._reset_envelope_filters
        ).pack(side="left", padx=(8, 0))

        self.env_story_var.trace_add("write", lambda *_: self._refresh_envelope_tree())
        self.env_pier_var.trace_add("write", lambda *_: self._refresh_envelope_tree())

        self.envelope_tree = self._create_treeview(self.tab_envelope)

    def _create_treeview(self, parent: ttk.Frame) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, show="headings", selectmode="extended")
        v_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self, padding=(10, 4, 10, 8))
        bar.pack(fill="x")

        ttk.Label(bar, textvariable=self.status_var).pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=160)
        self.progress.pack(side="right")

    # ------------------------------------------------------------------
    # State Helpers
    # ------------------------------------------------------------------
    def _set_busy(self, busy: bool, status_text: str | None = None) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn in (
            self.btn_connect,
            self.btn_pull,
            self.btn_recalc_top,
            self.btn_recalc_tab,
            self.btn_export,
        ):
            btn.configure(state=state)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
        if status_text is not None:
            self.status_var.set(status_text)

    def _run_in_background(
        self,
        status_text: str,
        worker: Callable[[], object],
        on_success: Callable[[object], None],
    ) -> None:
        if self._busy:
            return
        self._set_busy(True, status_text)

        def _target() -> None:
            try:
                result = worker()
            except Exception as exc:
                self.after(0, lambda err=exc: self._handle_worker_error(err))
                return
            self.after(0, lambda res=result: self._handle_worker_success(res, on_success))

        threading.Thread(target=_target, daemon=True).start()

    def _handle_worker_error(self, exc: Exception) -> None:
        self._set_busy(False, f"Error: {exc}")
        title = "ETABS Error" if isinstance(exc, ETABSError) else "Operation Failed"
        messagebox.showerror(title, str(exc), parent=self)

    def _handle_worker_success(
        self, result: object, on_success: Callable[[object], None]
    ) -> None:
        self._set_busy(False)
        on_success(result)

    def _populate_case_rows(
        self,
        load_cases: list[str],
        case_types: dict[str, str],
        existing_mappings: dict[str, str] | None = None,
    ) -> None:
        for widget in self.case_scroll.inner.winfo_children():
            widget.destroy()
        self._case_vars.clear()

        if not load_cases:
            ttk.Label(
                self.case_scroll.inner,
                text="No load cases found in the active ETABS model.",
                padding=12,
            ).pack(anchor="w")
            return

        mappings = existing_mappings or {}
        for case_name in load_cases:
            row = ttk.Frame(self.case_scroll.inner)
            row.pack(fill="x", padx=4, pady=2)

            is_selected = case_name in mappings
            default_cat = mappings.get(case_name, _suggest_nbcc_category(case_name))

            sel_var = tk.BooleanVar(value=is_selected)
            cat_var = tk.StringVar(value=default_cat)
            sel_var.trace_add("write", lambda *_: self._on_case_mapping_changed())
            cat_var.trace_add("write", lambda *_: self._on_case_mapping_changed())

            ttk.Checkbutton(row, variable=sel_var, width=8).pack(side="left")
            ttk.Label(row, text=case_name, width=34).pack(side="left")
            ttk.Label(
                row, text=case_types.get(case_name, ""), width=24
            ).pack(side="left")
            cb = ttk.Combobox(
                row,
                textvariable=cat_var,
                values=list(NBCC_CATEGORIES),
                state="readonly",
                width=16,
            )
            cb.pack(side="left")

            self._case_vars[case_name] = (sel_var, cat_var)

        self._update_stale_indicator()
        self._refresh_combo_defs_tree()

    def _on_case_mapping_changed(self) -> None:
        self._update_stale_indicator()
        self._refresh_combo_defs_tree()

    @staticmethod
    def _populate_listbox(
        lb: tk.Listbox,
        items: list[str],
        selected_items: list[str] | None,
    ) -> None:
        lb.delete(0, tk.END)
        for item in items:
            lb.insert(tk.END, item)
        if selected_items is None:
            lb.select_set(0, tk.END)
        else:
            wanted = set(selected_items)
            for idx, item in enumerate(items):
                if item in wanted:
                    lb.select_set(idx)

    @staticmethod
    def _select_all_listbox(lb: tk.Listbox) -> None:
        lb.select_set(0, tk.END)

    @staticmethod
    def _clear_listbox(lb: tk.Listbox) -> None:
        lb.selection_clear(0, tk.END)

    @staticmethod
    def _get_listbox_selection(lb: tk.Listbox) -> list[str] | None:
        total = lb.size()
        if total == 0:
            return None
        indices = lb.curselection()
        if len(indices) == total:
            return None  # All selected
        return [str(lb.get(i)) for i in indices]

    def _select_all_cases(self) -> None:
        for sel_var, _ in self._case_vars.values():
            sel_var.set(True)

    def _clear_all_cases(self) -> None:
        for sel_var, _ in self._case_vars.values():
            sel_var.set(False)

    def _auto_suggest_categories(self) -> None:
        for case_name, (_, cat_var) in self._case_vars.items():
            cat_var.set(_suggest_nbcc_category(case_name))

    def _collect_case_mappings(self) -> dict[str, str]:
        return {
            case_name: cat_var.get()
            for case_name, (sel_var, cat_var) in self._case_vars.items()
            if sel_var.get()
        }

    def _build_current_config(self) -> AppConfig:
        try:
            decimals = int(self.decimals_var.get())
        except (tk.TclError, ValueError) as exc:
            raise ValueError("Decimal places must be a non-negative integer.") from exc

        try:
            threshold = float(self.near_zero_threshold_var.get().strip())
        except ValueError as exc:
            raise ValueError("Near-zero threshold must be a valid number (e.g. 1e-4).") from exc

        cfg = AppConfig(
            units=self.units_var.get(),
            case_mappings=self._collect_case_mappings(),
            cleaning=CleaningOptions(
                selected_piers=self._get_listbox_selection(self.pier_listbox),
                selected_stories=self._get_listbox_selection(self.story_listbox),
                station_filter=self.station_var.get(),
                round_decimals=decimals,
                drop_near_zero=bool(self.drop_near_zero_var.get()),
                near_zero_threshold=threshold,
                remove_duplicates=bool(self.remove_duplicates_var.get()),
                aggregate_subcases=bool(self.aggregate_subcases_var.get()),
                zero_lateral_axial=bool(self.zero_lateral_axial_var.get()),
                zero_gravity_shear_moment=bool(self.zero_gravity_shear_moment_var.get()),
            ),
            nbcc=NBCCOptions(
                storage_assembly_occupancy=bool(self.storage_occupancy_var.get()),
                custom_combinations=(
                    [
                        ComboDefinition(
                            name=c.name,
                            factors=dict(c.factors),
                            enabled=c.enabled,
                            case_id=c.case_id,
                        )
                        for c in self._custom_combos
                    ]
                    if self._custom_combos is not None
                    else None
                ),
            ),
        )
        cfg.validate()
        return cfg

    def _apply_config_to_ui(self, cfg: AppConfig) -> None:
        self.units_var.set(cfg.units)
        self.station_var.set(cfg.cleaning.station_filter)
        self.decimals_var.set(cfg.cleaning.round_decimals)
        self.drop_near_zero_var.set(cfg.cleaning.drop_near_zero)
        self.near_zero_threshold_var.set(f"{cfg.cleaning.near_zero_threshold:g}")
        self.remove_duplicates_var.set(cfg.cleaning.remove_duplicates)
        self.aggregate_subcases_var.set(cfg.cleaning.aggregate_subcases)
        self.zero_lateral_axial_var.set(cfg.cleaning.zero_lateral_axial)
        self.zero_gravity_shear_moment_var.set(cfg.cleaning.zero_gravity_shear_moment)
        self.storage_occupancy_var.set(cfg.nbcc.storage_assembly_occupancy)
        self._custom_combos = (
            [
                ComboDefinition(
                    name=c.name,
                    factors=dict(c.factors),
                    enabled=c.enabled,
                    case_id=c.case_id,
                )
                for c in cfg.nbcc.custom_combinations
            ]
            if cfg.nbcc.custom_combinations is not None
            else None
        )

        if self._model_info is not None:
            self._populate_case_rows(
                self._model_info.load_cases,
                self._model_info.case_types,
                existing_mappings=cfg.case_mappings,
            )
            self._populate_listbox(
                self.pier_listbox,
                self._model_info.pier_labels,
                cfg.cleaning.selected_piers,
            )
            self._populate_listbox(
                self.story_listbox,
                self._model_info.stories,
                cfg.cleaning.selected_stories,
            )
        self._update_stale_indicator()
        self._refresh_combo_defs_tree()

    def _update_stale_indicator(self) -> None:
        if self._raw_df is None:
            self.stale_notice_var.set("")
            return

        selected_now = frozenset(self._collect_case_mappings().keys())
        reasons: list[str] = []
        if selected_now != self._pulled_cases:
            missing_from_cache = sorted(selected_now - self._pulled_cases)
            if missing_from_cache:
                reasons.append(
                    f"newly selected case(s) not in cache ({', '.join(missing_from_cache)})"
                )
            else:
                reasons.append("selected load case list changed")
        if self._pulled_units and self.units_var.get() != self._pulled_units:
            reasons.append(
                f"units changed ({self._pulled_units} -> {self.units_var.get()})"
            )

        if reasons:
            self.stale_notice_var.set(
                "Cache out of sync: "
                + "; ".join(reasons)
                + ". Click 'Pull from ETABS' to refresh raw data."
            )
        else:
            self.stale_notice_var.set("")

    # ------------------------------------------------------------------
    # Event Handlers
    # ------------------------------------------------------------------
    def _on_connect(self) -> None:
        units = self.units_var.get()
        self._run_in_background(
            "Connecting to ETABS and querying model metadata...",
            lambda: self._client.connect_and_query_metadata(units=units),
            self._on_connect_done,
        )

    def _on_connect_done(self, result: object) -> None:
        assert isinstance(result, ETABSModelInfo)
        self._model_info = result

        self.model_label_var.set(f"Model: {result.filename}")
        self.model_counts_var.set(
            f"Load Cases: {len(result.load_cases)} | "
            f"Piers: {len(result.pier_labels)} | "
            f"Stories: {len(result.stories)}"
        )

        # If user loaded a config before connecting, reconcile it now.
        if self._pending_config is not None:
            rec = reconcile_config_with_model(
                self._pending_config,
                result.load_cases,
                result.pier_labels,
                result.stories,
            )
            self._pending_config = None
            self._apply_config_to_ui(rec.config)
            if rec.has_warnings:
                messagebox.showwarning(
                    "Config Reconciled with Warnings",
                    "\n".join(rec.warnings),
                    parent=self,
                )
        else:
            existing = self._collect_case_mappings()
            self._populate_case_rows(
                result.load_cases, result.case_types, existing_mappings=existing
            )
            self._populate_listbox(self.pier_listbox, result.pier_labels, None)
            self._populate_listbox(self.story_listbox, result.stories, None)

        self.status_var.set(
            f"Connected to {result.filename} ({result.units}). "
            "Select load cases and click 'Pull from ETABS'."
        )

    def _on_pull(self) -> None:
        try:
            cfg = self._build_current_config()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc), parent=self)
            return

        selected_cases = list(cfg.case_mappings.keys())
        if not selected_cases:
            messagebox.showwarning(
                "No Load Cases Selected",
                "Select at least one ETABS load case in '1. Setup & Case Mapping' "
                "before pulling from ETABS.",
                parent=self,
            )
            self.notebook.select(self.tab_setup)
            return

        units = cfg.units
        self._run_in_background(
            f"Pulling pier forces for {len(selected_cases)} load case(s) from ETABS...",
            lambda: self._client.pull_pier_forces(selected_cases, units=units),
            lambda res: self._on_pull_done(res, frozenset(selected_cases)),
        )

    def _on_pull_done(
        self, result: object, pulled_cases: frozenset[str]
    ) -> None:
        assert isinstance(result, tuple)
        info, raw_df = result
        assert isinstance(info, ETABSModelInfo)
        assert isinstance(raw_df, pd.DataFrame)

        self._model_info = info
        self._raw_df = raw_df
        self._pulled_cases = pulled_cases
        self._pulled_units = info.units
        self._pulled_model_name = info.filename
        self.model_label_var.set(f"Model: {info.filename}")

        self._sync_raw_filter_options()
        self._refresh_raw_tree()
        self._update_stale_indicator()
        self._refresh_combo_defs_tree()
        self._run_calculation_from_cache(switch_to_envelope=True)

    def _on_recalculate(self) -> None:
        if self._raw_df is None:
            messagebox.showinfo(
                "No Cached Data",
                "No pier force data is cached yet. Click 'Pull from ETABS' first.",
                parent=self,
            )
            return
        self._run_calculation_from_cache(switch_to_envelope=False)

    def _run_calculation_from_cache(self, *, switch_to_envelope: bool) -> None:
        assert self._raw_df is not None
        try:
            cfg = self._build_current_config()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc), parent=self)
            return

        if not cfg.case_mappings:
            messagebox.showwarning(
                "No Load Cases Selected",
                "At least one load case must be checked to calculate combinations.",
                parent=self,
            )
            return

        story_order = self._model_info.stories if self._model_info else None
        pier_order = self._model_info.pier_labels if self._model_info else None

        try:
            # Keep both Top & Bottom through cleaning so NBCC superposition can pair
            # them to calculate Cm3 and Cm2 before applying the user's Station Filter.
            cleaned = clean_pier_forces(
                self._raw_df,
                cfg.case_mappings,
                cfg.cleaning,
                story_order=story_order,
                pier_order=pier_order,
                apply_station_filter=False,
            )
            combos_df, envelope_df = run_nbcc_analysis(
                cleaned, cfg.cleaning, cfg.nbcc
            )
        except Exception as exc:
            messagebox.showerror("Calculation Error", str(exc), parent=self)
            return

        self._combos_df = combos_df
        self._envelope_df = envelope_df

        self._sync_results_filter_options()
        self._refresh_combos_tree()
        self._refresh_envelope_tree()
        self._update_stale_indicator()

        self.status_var.set(
            f"Cached: {len(self._raw_df):,} raw rows from {self._pulled_model_name} "
            f"({self._pulled_units}) | {len(combos_df):,} combinations | "
            f"{len(envelope_df):,} envelope rows"
        )
        if switch_to_envelope and not envelope_df.empty:
            self.notebook.select(self.tab_envelope)

    def _on_save_config(self) -> None:
        try:
            cfg = self._build_current_config()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc), parent=self)
            return

        path = filedialog.asksaveasfilename(
            parent=self,
            title="Save Pier Forces Configuration",
            defaultextension=".json",
            filetypes=[("JSON Configuration", "*.json"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            save_config(cfg, path)
        except Exception as exc:
            messagebox.showerror("Save Failed", str(exc), parent=self)
            return
        self.status_var.set(f"Saved configuration to {Path(path).name}")

    def _on_load_config(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Load Pier Forces Configuration",
            filetypes=[("JSON Configuration", "*.json"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            loaded = load_config(path)
        except Exception as exc:
            messagebox.showerror("Load Failed", str(exc), parent=self)
            return

        if self._model_info is not None:
            rec = reconcile_config_with_model(
                loaded,
                self._model_info.load_cases,
                self._model_info.pier_labels,
                self._model_info.stories,
            )
            self._apply_config_to_ui(rec.config)
            if rec.has_warnings:
                messagebox.showwarning(
                    "Loaded Config with Warnings",
                    "\n".join(rec.warnings),
                    parent=self,
                )
            self.status_var.set(f"Loaded configuration from {Path(path).name}")
        else:
            self._pending_config = loaded
            self._apply_config_to_ui(loaded)
            self.status_var.set(
                f"Loaded {Path(path).name}. Connect to ETABS to match model items."
            )

    def _on_export_excel(self) -> None:
        if self._combos_df.empty and self._envelope_df.empty:
            messagebox.showinfo(
                "Nothing to Export",
                "Pull from ETABS and calculate combinations before exporting.",
                parent=self,
            )
            return

        try:
            cfg = self._build_current_config()
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc), parent=self)
            return

        stem = Path(self._pulled_model_name or "ETABS_Model").stem
        default_name = f"{stem}_Pier_Forces.xlsx"
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Export NBCC 2020 Pier Forces to Excel",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not path:
            return

        meta = ExportMetadata(
            model_filename=self._pulled_model_name or "Unknown",
            model_filepath=self._model_info.filepath if self._model_info else "",
            units=self._pulled_units or cfg.units,
            exported_at=datetime.now(),
            raw_row_count=len(self._raw_df) if self._raw_df is not None else 0,
            combo_row_count=len(self._combos_df),
            envelope_row_count=len(self._envelope_df),
        )

        try:
            export_to_excel(
                self._envelope_df,
                self._combos_df,
                cfg,
                meta,
                path,
            )
        except Exception as exc:
            messagebox.showerror("Export Failed", str(exc), parent=self)
            return

        self.status_var.set(f"Exported Excel workbook to {path}")
        messagebox.showinfo("Export Complete", f"Saved workbook:\n{path}", parent=self)

    def _on_copy_clipboard(self) -> None:
        active_tab = self.notebook.select()
        if active_tab == str(self.tab_envelope):
            df = self._get_filtered_envelope_df()
            label = "Envelope Summary"
            decimals = max(0, int(self.decimals_var.get()))
        elif active_tab == str(self.tab_combos):
            df = self._get_filtered_combos_df()
            label = "Load Combinations"
            decimals = max(0, int(self.decimals_var.get()))
        elif active_tab == str(self.tab_raw):
            df = self._get_filtered_raw_df()
            label = "Raw Imported Data"
            decimals = 4
        else:
            messagebox.showinfo(
                "Select a Data Tab",
                "Switch to '4. Raw Imported Data', '5. Load Combinations', or "
                "'6. Envelope Summary' to copy results to the clipboard.",
                parent=self,
            )
            return

        if df.empty:
            messagebox.showinfo(
                "No Rows to Copy",
                f"The {label} table is currently empty.",
                parent=self,
            )
            return

        tsv = dataframe_to_tsv(df, decimals=decimals)
        self.clipboard_clear()
        self.clipboard_append(tsv)
        self.update()
        self.status_var.set(f"Copied {len(df):,} {label} row(s) to clipboard (TSV).")

    # ------------------------------------------------------------------
    # Editable Combo Definitions Handlers
    # ------------------------------------------------------------------
    def _current_combo_defs(self) -> list[ComboDefinition]:
        if self._custom_combos is not None:
            return self._custom_combos
        mapped_cats = set(self._collect_case_mappings().values())
        nbcc_opts = NBCCOptions(
            storage_assembly_occupancy=bool(self.storage_occupancy_var.get())
        )
        return default_combo_definitions(mapped_cats, nbcc_opts)

    def _ensure_custom_combos_mutable(self) -> list[ComboDefinition]:
        if self._custom_combos is None:
            self._custom_combos = [
                ComboDefinition(
                    name=c.name,
                    factors=dict(c.factors),
                    enabled=c.enabled,
                    case_id=c.case_id,
                )
                for c in self._current_combo_defs()
            ]
        return self._custom_combos

    def _refresh_combo_defs_tree(self) -> None:
        if not hasattr(self, "combo_defs_tree"):
            return
        combos = self._current_combo_defs()
        enabled_count = sum(1 for c in combos if c.enabled)
        if self._custom_combos is None:
            self.combo_defs_mode_var.set(
                f"Mode: NBCC 2020 Table 4.1.3.2.-A Defaults ({len(combos)} combo(s))"
            )
        else:
            self.combo_defs_mode_var.set(
                f"Mode: Custom / Edited ({enabled_count} of {len(combos)} enabled)"
            )

        tree = self.combo_defs_tree
        tree.delete(*tree.get_children())
        for idx, combo in enumerate(combos):
            row_vals = [
                "Yes" if combo.enabled else "No",
                combo.case_id,
                combo.name,
                *[
                    f"{combo.factors[cat]:g}" if cat in combo.factors else ""
                    for cat in NBCC_CATEGORIES
                ],
            ]
            tree.insert("", tk.END, iid=str(idx), values=row_vals)

    def _on_add_combo_def(self) -> None:
        dlg = _ComboEditorDialog(self, title="Add Load Combination")
        self.wait_window(dlg)
        if dlg.result is None:
            return
        combos = self._ensure_custom_combos_mutable()
        combos.append(dlg.result)
        self._refresh_combo_defs_tree()

    def _on_edit_combo_def(self) -> None:
        selection = self.combo_defs_tree.selection()
        if not selection:
            messagebox.showinfo(
                "Select a Combination",
                "Select a combination row to edit.",
                parent=self,
            )
            return
        idx = int(selection[0])
        combos = self._current_combo_defs()
        if idx < 0 or idx >= len(combos):
            return
        dlg = _ComboEditorDialog(
            self, title="Edit Load Combination", initial=combos[idx]
        )
        self.wait_window(dlg)
        if dlg.result is None:
            return
        mutable = self._ensure_custom_combos_mutable()
        mutable[idx] = dlg.result
        self._refresh_combo_defs_tree()
        self.combo_defs_tree.selection_set(str(idx))

    def _on_toggle_combo_def(self) -> None:
        selection = self.combo_defs_tree.selection()
        if not selection:
            return
        mutable = self._ensure_custom_combos_mutable()
        indices = [int(i) for i in selection if 0 <= int(i) < len(mutable)]
        for idx in indices:
            mutable[idx].enabled = not mutable[idx].enabled
        self._refresh_combo_defs_tree()
        for idx in indices:
            self.combo_defs_tree.selection_add(str(idx))

    def _on_delete_combo_def(self) -> None:
        selection = self.combo_defs_tree.selection()
        if not selection:
            return
        mutable = self._ensure_custom_combos_mutable()
        indices = sorted((int(i) for i in selection), reverse=True)
        for idx in indices:
            if 0 <= idx < len(mutable):
                mutable.pop(idx)
        self._refresh_combo_defs_tree()

    def _on_reset_combo_defs(self) -> None:
        self._custom_combos = None
        self._refresh_combo_defs_tree()
        self.status_var.set("Reset load combination definitions to NBCC 2020 defaults.")

    # ------------------------------------------------------------------
    # Raw Imported Data Filters & Display
    # ------------------------------------------------------------------
    def _reset_raw_filters(self) -> None:
        self.raw_story_var.set("All Stories")
        self.raw_pier_var.set("All Piers")
        self.raw_case_var.set("All Cases")
        self.raw_loc_var.set("All")

    def _sync_raw_filter_options(self) -> None:
        df = self._raw_df
        if df is None or df.empty:
            stories = ["All Stories"]
            piers = ["All Piers"]
            cases = ["All Cases"]
        else:
            stories = ["All Stories"] + list(dict.fromkeys(df["Story"].astype(str)))
            piers = ["All Piers"] + list(dict.fromkeys(df["Pier"].astype(str)))
            cases = ["All Cases"] + list(dict.fromkeys(df["LoadCase"].astype(str)))

        self.raw_story_cb.configure(values=stories)
        self.raw_pier_cb.configure(values=piers)
        self.raw_case_cb.configure(values=cases)
        if self.raw_story_var.get() not in stories:
            self.raw_story_var.set("All Stories")
        if self.raw_pier_var.get() not in piers:
            self.raw_pier_var.set("All Piers")
        if self.raw_case_var.get() not in cases:
            self.raw_case_var.set("All Cases")

    def _get_filtered_raw_df(self) -> pd.DataFrame:
        if self._raw_df is None or self._raw_df.empty:
            return pd.DataFrame()
        df = self._raw_df
        mask = pd.Series(True, index=df.index)
        story = self.raw_story_var.get()
        if story and story != "All Stories":
            mask &= df["Story"].astype(str) == story
        pier = self.raw_pier_var.get()
        if pier and pier != "All Piers":
            mask &= df["Pier"].astype(str) == pier
        case = self.raw_case_var.get()
        if case and case != "All Cases":
            mask &= df["LoadCase"].astype(str) == case
        loc = self.raw_loc_var.get()
        if loc and loc != "All":
            mask &= df["Location"].astype(str) == loc
        return df.loc[mask]

    def _refresh_raw_tree(self) -> None:
        if not hasattr(self, "raw_tree"):
            return
        df = self._get_filtered_raw_df()
        total = len(self._raw_df) if self._raw_df is not None else 0
        if len(df) == total:
            self.raw_count_var.set(f"{total:,} raw imported row(s) (uncleaned)")
        else:
            self.raw_count_var.set(
                f"Showing {len(df):,} of {total:,} raw imported row(s) (uncleaned)"
            )
        self._populate_treeview(self.raw_tree, df, decimals_override=4)

    # ------------------------------------------------------------------
    # Treeview Population & View Filters
    # ------------------------------------------------------------------
    def _reset_combo_filters(self) -> None:
        self.combo_story_var.set("All Stories")
        self.combo_pier_var.set("All Piers")
        self.combo_search_var.set("")

    def _reset_envelope_filters(self) -> None:
        self.env_story_var.set("All Stories")
        self.env_pier_var.set("All Piers")

    def _sync_results_filter_options(self) -> None:
        combo_stories = (
            ["All Stories"] + list(dict.fromkeys(self._combos_df["Story"].astype(str)))
            if not self._combos_df.empty
            else ["All Stories"]
        )
        combo_piers = (
            ["All Piers"] + list(dict.fromkeys(self._combos_df["Pier"].astype(str)))
            if not self._combos_df.empty
            else ["All Piers"]
        )
        self.combo_story_cb.configure(values=combo_stories)
        self.combo_pier_cb.configure(values=combo_piers)
        if self.combo_story_var.get() not in combo_stories:
            self.combo_story_var.set("All Stories")
        if self.combo_pier_var.get() not in combo_piers:
            self.combo_pier_var.set("All Piers")

        env_stories = (
            ["All Stories"] + list(dict.fromkeys(self._envelope_df["Story"].astype(str)))
            if not self._envelope_df.empty
            else ["All Stories"]
        )
        env_piers = (
            ["All Piers"] + list(dict.fromkeys(self._envelope_df["Pier"].astype(str)))
            if not self._envelope_df.empty
            else ["All Piers"]
        )
        self.env_story_cb.configure(values=env_stories)
        self.env_pier_cb.configure(values=env_piers)
        if self.env_story_var.get() not in env_stories:
            self.env_story_var.set("All Stories")
        if self.env_pier_var.get() not in env_piers:
            self.env_pier_var.set("All Piers")

    def _get_filtered_combos_df(self) -> pd.DataFrame:
        df = self._combos_df
        if df.empty:
            return df
        mask = pd.Series(True, index=df.index)
        story = self.combo_story_var.get()
        if story and story != "All Stories":
            mask &= df["Story"].astype(str) == story
        pier = self.combo_pier_var.get()
        if pier and pier != "All Piers":
            mask &= df["Pier"].astype(str) == pier
        query = self.combo_search_var.get().strip().lower()
        if query:
            mask &= (
                df["Story"].astype(str).str.lower().str.contains(query, regex=False)
                | df["Pier"].astype(str).str.lower().str.contains(query, regex=False)
                | df["ComboName"].astype(str).str.lower().str.contains(query, regex=False)
            )
        return df.loc[mask]

    def _get_filtered_envelope_df(self) -> pd.DataFrame:
        df = self._envelope_df
        if df.empty:
            return df
        mask = pd.Series(True, index=df.index)
        story = self.env_story_var.get()
        if story and story != "All Stories":
            mask &= df["Story"].astype(str) == story
        pier = self.env_pier_var.get()
        if pier and pier != "All Piers":
            mask &= df["Pier"].astype(str) == pier
        return df.loc[mask]

    def _refresh_combos_tree(self) -> None:
        df = self._get_filtered_combos_df()
        total = len(self._combos_df)
        if len(df) == total:
            self.combo_count_var.set(f"{total:,} load combination row(s)")
        else:
            self.combo_count_var.set(
                f"Showing {len(df):,} of {total:,} load combination row(s)"
            )
        self._populate_treeview(self.combos_tree, df)

    def _refresh_envelope_tree(self) -> None:
        df = self._get_filtered_envelope_df()
        total = len(self._envelope_df)
        if len(df) == total:
            self.envelope_count_var.set(f"{total:,} governing envelope row(s)")
        else:
            self.envelope_count_var.set(
                f"Showing {len(df):,} of {total:,} governing envelope row(s)"
            )
        self._populate_treeview(self.envelope_tree, df)

    def _populate_treeview(
        self,
        tree: ttk.Treeview,
        df: pd.DataFrame,
        *,
        decimals_override: int | None = None,
    ) -> None:
        tree.delete(*tree.get_children())
        cols = list(df.columns)
        tree["columns"] = cols
        if not cols:
            return

        if decimals_override is not None:
            decimals = decimals_override
        else:
            try:
                decimals = max(0, int(self.decimals_var.get()))
            except (tk.TclError, ValueError):
                decimals = 2
        num_fmt = f"{{:,.{decimals}f}}"
        cm_fmt = f"{{:,.{max(2, decimals)}f}}"

        numeric_cols = {
            col for col in cols if pd.api.types.is_numeric_dtype(df[col])
        }
        cm_cols = {"Cm3", "Cm2"}

        for col in cols:
            is_num = col in numeric_cols
            anchor = "e" if is_num else "w"
            if "Combo" in col:
                width = 220
            elif col in cm_cols:
                width = 72
            elif is_num:
                width = 92
            else:
                width = 88
            tree.heading(col, text=col, anchor="center")
            tree.column(col, width=width, minwidth=64, anchor=anchor, stretch=False)

        for row_vals in df.itertuples(index=False, name=None):
            display = []
            for col, val in zip(cols, row_vals):
                if pd.isna(val):
                    display.append("")
                elif col in numeric_cols:
                    num = float(val)
                    if num == 0.0:
                        num = 0.0
                    display.append((cm_fmt if col in cm_cols else num_fmt).format(num))
                else:
                    display.append(str(val))
            tree.insert("", tk.END, values=display)


def launch_app() -> None:
    """Create and run the Tkinter desktop application."""
    app = PierForcesApp()
    app.mainloop()
