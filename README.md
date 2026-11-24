# ETABS Pier Forces (NBCC 2020)

A Windows desktop application (`tkinter`/`ttk`, `pandas`, `openpyxl`, `comtypes`) that connects directly to an open **CSI ETABS** model, extracts wall **Pier Forces** for user-selected load cases, cleans and aggregates the data, computes **NBCC 2020 Table 4.1.3.2.-A** ULS load combinations with **CSA A23.3** equivalent uniform moment factors ($C_m$), generates a concurrent-force **Envelope Summary**, and exports formatted Excel workbooks.

---

## Table of Contents

- [Features](#features)
- [Supported Software & Requirements](#supported-software--requirements)
- [Installation & Setup with `uv`](#installation--setup-with-uv)
- [Running the Application & Tests](#running-the-application--tests)
- [How It Works (Engineering & Code Reference)](#how-it-works-engineering--code-reference)
  - [NBCC 2020 Table 4.1.3.2.-A Load Combinations](#nbcc-2020-table-4132-a-load-combinations)
  - [CSA A23.3 Equivalent Uniform Moment Factors (`Cm3` & `Cm2`)](#csa-a233-equivalent-uniform-moment-factors-cm3--cm2)
  - [Multi-Step Seismic & Directional Lateral Cases](#multi-step-seismic--directional-lateral-cases)
  - [Units, Sign Conventions & Precision](#units-sign-conventions--precision)
- [User Guide](#user-guide)
  - [Top Toolbar](#top-toolbar)
  - [Tab 1: Setup & Case Mapping](#tab-1-setup--case-mapping)
  - [Tab 2: Cleaning & NBCC Options](#tab-2-cleaning--nbcc-options)
  - [Tab 3: Combo Definitions](#tab-3-combo-definitions)
  - [Tab 4: Raw Imported Data](#tab-4-raw-imported-data)
  - [Tab 5: Load Combinations](#tab-5-load-combinations)
  - [Tab 6: Envelope Summary](#tab-6-envelope-summary)
  - [Excel Workbook Export Structure](#excel-workbook-export-structure)
  - [JSON Project Presets](#json-project-presets)
- [Project Architecture](#project-architecture)
- [Troubleshooting](#troubleshooting)

---

## Features

- **Direct ETABS COM Integration**: Attaches to an already-open ETABS instance via `CSI.ETABS.API.ETABSObject` without hardcoded DLL paths or `.NET` version lock-in.
- **Fast Selective Extraction**: Queries available Load Cases, Pier Labels, and Story Names upfront, then selects only the user's chosen load cases in ETABS before calling `SapModel.Results.PierForce`.
- **Two-Step Cached Workflow**: Pull raw forces once in a background thread (UI never freezes), then recalculate instantaneously in memory as you change pier/story/station filters, cleaning toggles, occupancy settings, or custom load combinations.
- **NBCC 2020 Table 4.1.3.2.-A Engine**: Declarative, line-by-line auditable ULS superposition supporting standard and Storage/Assembly occupancies, plus a full **Combo Definitions** editor for custom or modified combinations.
- **Strong- & Weak-Axis $C_m$ Factors**: Automatically pairs `Top` and `Bottom` factored moments for every combination to calculate `Cm3` and `Cm2` per **CSA A23.3 Clause 10.15.3.1**, even when outputting `Bottom`-only forces for wall design.
- **Concurrent-Force Envelope Summary**: Produces 10 governing extreme rows (`P`, `V2`, `V3`, `M2`, `M3` Min & Max) per `(Pier, Story, Location)`, preserving all concurrent forces and $C_m$ factors from the winning load combination for direct shear wall P–M interaction checks.
- **Raw Data Audit & Level Filtering**: Inspect uncleaned ETABS extractions on a dedicated **Raw Imported Data** tab and filter results by **Level (Story)**, **Pier**, **Load Case**, or **Location**.
- **Formatted Excel & Clipboard Export**: Generates a 3-sheet `.xlsx` report (`Envelope Summary`, `Load Combinations`, `Metadata`) using native Excel Tables with auto-fitted widths, frozen headers, and an audit trail—plus one-click tab-separated clipboard copy.
- **Reusable JSON Presets**: Save and reload case mappings, filters, cleaning options, and custom load combinations across models with non-blocking warnings for any missing model items.

---

## Supported Software & Requirements

| Requirement | Details |
| :--- | :--- |
| **Operating System** | Windows 10 / 11 (required for CSI ETABS COM automation) |
| **CSI ETABS** | **ETABS v18 – v23+** (any 64-bit Windows version supporting the standard OAPI COM ProgID `CSI.ETABS.API.ETABSObject` and `SapModel.Results.PierForce`, `PierLabel`, `Story`, `LoadCases`, and `SetPresentUnits`). Verified for modern ETABS v22 / v23 workflows. |
| **Python** | Python **3.13+** (automatically managed by `uv`) |
| **Package Manager** | [`uv`](https://docs.astral.sh/uv/) (exclusively manages virtual environments, `pyproject.toml`, and `uv.lock`) |

> **Note:** An active ETABS installation is only required to connect and pull live model data. All downstream calculation, cleaning, configuration, export, and unit-test modules run independently without ETABS installed or open.

---

## Installation & Setup with `uv`

### 1. Install `uv` (if not already installed)

In PowerShell:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Or via `winget`:

```powershell
winget install --id=astral-sh.uv -e
```

Restart your terminal and confirm `uv` is available:

```powershell
uv --version
```

### 2. Clone the Repository & Sync Dependencies

```powershell
git clone <your-repo-url>
cd PierForces
uv sync
```

`uv sync` automatically creates `.venv`, installs Python 3.13 (if needed), and installs exact locked versions of `pandas`, `openpyxl`, `comtypes`, and `pytest` from `uv.lock`.

---

## Running the Application & Tests

### Launch the Desktop Application

1. Open your model in **CSI ETABS** and make sure analysis has been run for the load cases you want to extract.
2. Start the app via `uv`:

```powershell
uv run main.py
```

### Run the Automated Test Suite (No ETABS Required)

```powershell
uv run pytest -v
```

### Verify Lockfile Integrity

```powershell
uv lock --check
```

---

## How It Works (Engineering & Code Reference)

### NBCC 2020 Table 4.1.3.2.-A Load Combinations

By default, the engine generates ULS combinations from **NBCC 2020, Division B, Part 4, Table 4.1.3.2.-A** (*Load Combinations Without Crane Loads for Ultimate Limit States*), emitting only combinations whose required categories are mapped by the user:

| NBCC Case | Principal Load(s) | Companion Load(s) | Notes |
| :---: | :--- | :--- | :--- |
| **Case 1** | `1.4D` | — | Pure dead load |
| **Case 2** | `(1.25D or 0.9D) + 1.5L` | `1.0S` or `±0.4W` | Live principal; base + companion variants |
| **Case 3** | `(1.25D or 0.9D) + 1.5S` | `1.0L` or `±0.4W` | Snow principal; companion `L` is always `1.0L` per NBCC 2020 |
| **Case 4** | `(1.25D or 0.9D) ± 1.4W` | `0.5L` *(or `1.0L`)* or `0.5S` | Wind principal with `±` reversals |
| **Case 5** | `1.0D ± 1.0E` | `0.5L` *(or `1.0L`)* `+ 0.25S` | Earthquake principal with `±` reversals; base + companion variants |

- **Companion-Omitted Variants (NBCC 2020 Clause 4.1.3.2.(2))**: Because companion loads are omitted when they reduce a load effect (e.g., overturning or tension), both principal-only combinations (`1.25D + 1.5L`, `0.9D ± 1.4W`, `1.0D ± 1.0E`) and companion combinations are evaluated whenever companion loads are present.
- **Storage / Assembly Occupancy Toggle (NBCC 2020 Clause 4.1.3.2.(7) / Table 4.1.3.2.-A Note 6)**: Checking this option increases the companion live-load factor from **`0.5L` → `1.0L`** in **Cases 4 and 5** (Case 3 remains `1.0L` for all occupancies).
- **Custom / Edited Combinations**: All combinations can be inspected, enabled/disabled, modified, or supplemented in **Tab 3 (`3. Combo Definitions`)**, and reset to strict NBCC 2020 defaults at any time.

### CSA A23.3 Equivalent Uniform Moment Factors (`Cm3` & `Cm2`)

For wall/pier slenderness design, the app computes both strong-axis (**`Cm3`** from `M3`) and weak-axis (**`Cm2`** from `M2`) factors per **CSA A23.3 Clause 10.15.3.1**:

$$C_m = \max\!\left(0.4,\; 0.6 + 0.4\,\frac{M_1}{M_2}\right)$$

where $|M_1| \le |M_2|$ are the factored end moments at the `Top` and `Bottom` of the pier for the same `(Story, Pier, ComboName)`.

- **ETABS Curvature Sign Convention**:
  - **Same sign** at `Top` and `Bottom` (`M_top * M_bot > 0`) → **Single curvature** ($M_1 / M_2 > 0$, yielding $0.6 < C_m \le 1.0$).
  - **Opposite signs** at `Top` and `Bottom` (`M_top * M_bot < 0`) → **Double curvature** ($M_1 / M_2 < 0$, yielding $0.4 \le C_m < 0.6$).
  - **Zero moment at one end** → $M_1/M_2 = 0 \Rightarrow C_m = 0.6$.
  - **Both end moments effectively zero** ($\le 10^{-12}$) or opposite station unavailable → conservative default **`1.0`**.
- **Design at `Bottom` with `Top` Moments Preserved**: Even when **Station Filter** is set to **`Bottom`** (the default), the engine internally runs NBCC superposition for **both `Top` and `Bottom`** so `Cm3` and `Cm2` are calculated from the true paired end moments before filtering the final displayed and exported tables to `Bottom`.

### Multi-Step Seismic & Directional Lateral Cases

- **Multi-Step Preservation**: When ETABS returns multiple rows for the same `(Story, Pier, LoadCase, Location)` (such as static seismic cases with accidental eccentricity steps), the cleaner automatically labels them `<CaseName>-1`, `<CaseName>-2`, `<CaseName>-3` and expands each step into its own separate load combination—preventing steps from ever being summed together.
- **Gravity Aggregation vs. Directional Lateral Separation**:
  - Multiple gravity sub-cases mapped to the same load type (e.g., `DL_Self` + `SDL` → `Dead Load`) are summed into a single category total when **Aggregate gravity sub-cases** is enabled.
  - Distinct directional `Wind Load` (`WX`, `WY`) and `Seismic Load` (`EQx`, `EQy`) cases are **always kept separate** and combined individually with `±` sign reversals.

### Units, Sign Conventions & Precision

- **Output Units**: Defaults to **`kN, m, C`**, with a selector supporting `kN, m, C`, `N, mm, C`, `kip, ft, F`, and `kip, in, F`.
- **Sign Convention**: Preserves native ETABS pier force signs without modification (negative `P` = compression, positive `P` = tension).
- **Full-Precision Math**: Sub-case aggregation, load-factor superposition, $C_m$ calculation, and governing envelope selection are all performed at full 64-bit floating-point precision. User decimal rounding (default: `2`) is applied only to the final displayed/exported outputs, and `Cm3`/`Cm2` always retain at least 2 decimal places (even if force rounding is set to `0`).

---

## User Guide

### Top Toolbar

| Button | Action |
| :--- | :--- |
| **Connect to ETABS** | Attaches to the open ETABS model in a background thread, sets output units, and populates Load Cases, Pier Labels, and Stories. |
| **Pull from ETABS** | Selects only your checked load cases in ETABS, extracts raw `PierForce` rows in a background thread, caches them in memory, populates **Raw Imported Data**, and runs the initial calculation. |
| **Recalculate** | Re-runs cleaning, combinations, $C_m$, and envelope calculations instantaneously against the in-memory raw cache—no extra COM call to ETABS. |
| **Save Config** | Saves your unit choice, case mappings, filters, cleaning options, occupancy toggle, and any custom combinations to a `.json` preset file. |
| **Load Config** | Loads a `.json` preset. When connected to a model, automatically matches existing cases/piers/stories and displays a non-blocking warning for any items not present in the active model. |
| **Export to Excel (.xlsx)** | Writes the formatted 3-sheet workbook (`Envelope Summary`, `Load Combinations`, `Metadata`). |
| **Copy to Clipboard** | Copies the currently active data tab (`4. Raw Imported Data`, `5. Load Combinations`, or `6. Envelope Summary`), respecting active view filters, as tab-separated text ready to paste directly into Excel. |

> **Stale-Cache Indicator**: If you check/uncheck load cases or change output units after pulling, an amber warning reminds you to click **Pull from ETABS** so the raw cache stays in sync.

### Tab 1: Setup & Case Mapping

1. Choose **Output Units** (default `kN, m, C`) and click **Connect to ETABS**.
2. Check the ETABS load cases you want to extract and assign each one an NBCC load type (`Dead Load`, `Live Load`, `Snow Load`, `Wind Load`, or `Seismic Load`).
3. Use **Auto-Suggest Load Types** if you want quick heuristic suggestions based on common case names (`DEAD`, `SDL`, `LIVE`, `SNOW`, `WX`, `EQx`, `RSX`, etc.), then verify each mapping.
4. Click **Pull from ETABS** to extract and cache results.

### Tab 2: Cleaning & NBCC Options

- **Pier Label Filter** & **Story Filter**: Multi-select lists populated upfront from ETABS (with `Select All` / `Clear All`; defaults to all piers and all stories in model order).
- **Station Filter**: `Bottom` (default), `Top`, or `All`. (`Top` and `Bottom` are still paired internally to compute `Cm3`/`Cm2`.)
- **Decimal Places**: Controls output rounding and Excel number formatting (default `2`; `Cm3`/`Cm2` always keep at least 2 decimals).
- **Drop Near-Zero Rows**: Optionally drops cleaned rows where all 6 force/moment magnitudes are below the threshold (default `1e-4`).
- **Remove Exact Duplicate Rows**: Drops identical duplicate rows without disturbing distinct multi-step seismic rows.
- **Aggregate Gravity Sub-Cases (Dead, Live, Snow)**: Enabled by default to sum multiple `Dead Load`, `Live Load`, or `Snow Load` sub-cases before combining.
- **Zero Axial Force (P) for Lateral Cases (Wind, Seismic)** *(unchecked by default)*: Optional idealized-wall cleanup that zeroes lateral `P`.
- **Zero Shear & Moments for Gravity Cases (Dead, Live, Snow)** *(unchecked by default)*: Optional idealized-wall cleanup that zeroes gravity `V2`, `V3`, `T`, `M2`, and `M3`.
- **Storage / Assembly Occupancy**: Switches companion live load from `0.5L` to `1.0L` in NBCC Cases 4 and 5.
- Click **Recalculate from Cached Data** after changing any option.

### Tab 3: Combo Definitions

Displays every active load combination definition (`Enabled | Case | Combo Name | Dead Load | Live Load | Snow Load | Wind Load | Seismic Load`).

- **Add Combo**: Create a custom combination (e.g., serviceability `1.0D + 1.0L` or custom factors) with an **Auto-Name** helper.
- **Edit Selected** *(or double-click a row)*: Modify the name, enabled state, or any category factor.
- **Toggle Enable** *(or press `Space`)*: Quickly include/exclude combinations without deleting them.
- **Delete Selected**: Remove unwanted combinations.
- **Reset to NBCC 2020 Defaults**: Discard custom edits and regenerate strict NBCC 2020 Table 4.1.3.2.-A definitions from your active case mappings and occupancy toggle.
- **Apply & Recalculate**: Immediately recompute all results from the cached raw pull.

### Tab 4: Raw Imported Data

Displays the exact, uncleaned `PierForce` rows pulled from ETABS (`Story | Pier | LoadCase | Location | P | V2 | V3 | T | M2 | M3`) at 4-decimal display precision before any filtering, aggregation, zeroing, or superposition. Filter interactively by **Level (Story)**, **Pier**, **Load Case**, and **Location**.

### Tab 5: Load Combinations

Displays all calculated factored load combinations (`Story | Pier | Location | ComboName | P | T | V2 | M3 | Cm3 | V3 | M2 | Cm2`) with interactive **Level (Story)** and **Pier** dropdowns plus free-text **Search**.

### Tab 6: Envelope Summary

Displays the governing concurrent-force envelope grouped by `(Pier, Story, Location)` with 10 extreme rows per group (`P Min`, `P Max`, `V2 Min`, `V2 Max`, `V3 Min`, `V3 Max`, `M2 Min`, `M2 Max`, `M3 Min`, `M3 Max`) and exact column order:

```text
Pier | Story | Combo | Extreme | Location | P | T | V2 | M3 | Cm3 | V3 | M2 | Cm2
```

Includes interactive **Level (Story)** and **Pier** dropdowns for quick review and filtered clipboard copying.

### Excel Workbook Export Structure

Exporting produces a single `.xlsx` file containing three worksheets:

1. **`Envelope Summary`**: Native Excel Table (`Envelope_Summary`, `TableStyleMedium2`) with blue headers, banded rows, filter dropdowns, frozen top row (`A2`), auto-fitted widths, and dynamic number formatting.
2. **`Load Combinations`**: Native Excel Table (`Load_Combinations`, `TableStyleMedium2`) with all factored combinations and `Cm3`/`Cm2` factors.
3. **`Metadata`**: Structured engineering audit trail logging:
   - Model filename, full path, export timestamp, active units, and row counts
   - ETABS Load Case → NBCC Category mapping table
   - All cleaning filters, tolerances, zeroing toggles, and occupancy settings
   - Load Combination Mode (`NBCC 2020 Table 4.1.3.2.-A (Default)` vs. `Custom / Edited`) and the full list of active combination formulas

### JSON Project Presets

Example saved `.json` preset:

```json
{
  "schema_version": 1,
  "units": "kN, m, C",
  "case_mappings": {
    "SelfWeight": "Dead Load",
    "SDL": "Dead Load",
    "Live": "Live Load",
    "Snow": "Snow Load",
    "EQx": "Seismic Load",
    "EQy": "Seismic Load"
  },
  "cleaning": {
    "selected_piers": null,
    "selected_stories": null,
    "station_filter": "Bottom",
    "round_decimals": 2,
    "drop_near_zero": false,
    "near_zero_threshold": 0.0001,
    "remove_duplicates": false,
    "aggregate_subcases": true,
    "zero_lateral_axial": false,
    "zero_gravity_shear_moment": false
  },
  "nbcc": {
    "storage_assembly_occupancy": false,
    "custom_combinations": null
  }
}
```

---

## Project Architecture

All `comtypes` / ETABS COM logic is strictly isolated in `src/etabs_client.py`. All downstream modules operate on pure `pandas` DataFrames and Python dataclasses so they remain 100% testable without ETABS running.

```text
PierForces/
├── main.py                    # Application entry point
├── pyproject.toml             # uv project configuration and dependencies
├── uv.lock                    # Locked dependency tree
├── src/
│   ├── etabs_client.py        # Isolated comtypes ETABS OAPI client (thread-safe COM)
│   ├── config.py              # Dataclasses, JSON preset save/load, model reconciliation
│   ├── cleaner.py             # Filtering, multi-step preservation, gravity aggregation, cleanup
│   ├── nbcc_combos.py         # Declarative NBCC 2020 Table 4.1.3.2.-A, CSA A23.3 Cm, envelope
│   ├── exporter.py            # 3-sheet openpyxl Excel Table export & TSV clipboard formatter
│   └── ui.py                  # 6-tab Tkinter/ttk GUI, raw cache, background worker threads
└── tests/
    ├── test_config.py         # Config validation, JSON roundtrip, reconciliation tests
    ├── test_cleaner.py        # Filtering, multi-step separation, aggregation, zeroing tests
    ├── test_nbcc_combos.py    # Hand-verified NBCC 2020 Cases 1–5, Cm, custom combos, envelope
    ├── test_exporter.py       # 3-sheet Excel Table structure, formatting, Metadata, TSV tests
    ├── test_etabs_client.py   # Mocked OAPI extraction and error-handling tests
    └── test_end_to_end.py     # Full pipeline + Tkinter cache/filter/editor smoke tests
```

---

## Troubleshooting

- **"Could not attach to an open CSI ETABS instance"**:
  - Ensure CSI ETABS is open with a model file loaded.
  - **Windows UAC / Privilege Mismatch**: If ETABS was launched *As Administrator*, Windows COM blocks non-elevated processes from attaching via `GetActiveObject`. Run ETABS and your terminal at the same privilege level (normally both non-elevated).
- **"ETABS returned 0 pier force rows"**:
  - Confirm the model has been **analyzed** for the selected load cases and that **Pier Labels** are assigned to wall shells.
- **Selected load cases changed after pulling**:
  - Changes to cleaning filters, rounding, occupancy, or combo definitions only require clicking **Recalculate**. Adding new ETABS load cases or changing output units requires clicking **Pull from ETABS** to refresh the cached raw data.
