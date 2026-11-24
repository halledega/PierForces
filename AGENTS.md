# AGENTS.md — ETABS Pier Forces (NBCC 2020)

## Project Overview
A Python desktop application (`tkinter`/`ttk`, `pandas`, `openpyxl`, `comtypes`) managed with `uv` that extracts Pier Forces from an open CSI ETABS model, cleans the data, computes NBCC 2020 ULS load combinations and governing envelopes, and exports formatted multi-sheet Excel workbooks.

---

## Core Agent Directives

### 1. No Assumptions — Ask When Ambiguous
- Never guess or assume requirements, structural engineering conventions, edge cases, or UI/UX behavior.
- Whenever anything is unclear or ambiguous, stop and ask any and all questions needed before writing code.
- You may provide concise, concrete recommendations alongside your questions, but you must wait for explicit user confirmation before acting on them.

### 2. Read-Only Reference & Plan-First Gate
- Treat `C:\Users\MichaelHalliday\Documents\_Repos\PyCSI` as **strictly read-only**. Never modify, overwrite, or create files inside that directory.
- Before implementing new features or refactoring reference logic, inspect the relevant files, list any bugs or discrepancies found, and present a clear implementation plan for approval.

### 3. Strict `uv` Environment Discipline
- Use `uv` exclusively for all environment, dependency, and execution commands (`uv init`, `uv add`, `uv run`, `uv run pytest`).
- Never invoke `pip`, `pip install`, or bare global `python` commands directly.
- Keep `pyproject.toml` and `uv.lock` clean and in sync.

### 4. Structural Engineering & Code-Compliance Rigor
- Never invent, simplify, or approximate NBCC 2020 load factors, companion factors, or sign conventions.
- If the reference notebook in `PyCSI` and **NBCC 2020 Table 4.1.3.2.-A** conflict, flag the discrepancy explicitly and ask how to proceed before coding.
- Keep all load combination definitions in `src/nbcc_combos.py` declarative, transparent, and easy to audit line-by-line.

### 5. Separation of Concerns (COM Isolation)
- Isolate all `comtypes` and ETABS OAPI calls strictly inside `src/etabs_client.py`.
- All downstream modules (`src/cleaner.py`, `src/nbcc_combos.py`, `src/config.py`, `src/exporter.py`) must accept and return pure `pandas` DataFrames or standard Python dataclasses/dicts so they remain 100% testable without ETABS running.
- Never block the Tkinter main thread with COM calls; always initialize COM (`comtypes.CoInitialize()`) inside any worker thread that touches ETABS.

### 6. Incremental Verification
- Build and verify one module at a time rather than generating the entire application in one unchecked pass.
- After implementing or modifying `src/cleaner.py`, `src/nbcc_combos.py`, `src/config.py`, or `src/exporter.py`, run `uv run pytest` (or a targeted verification script) to confirm behavior.
- Do not mark a task complete until tests pass cleanly without errors.

---

## Architecture Boundaries
- `main.py` — Entry point only.
- `src/etabs_client.py` — Active ETABS connection, units, metadata queries (`LoadCases`, `PierLabel`, `Story`), and `PierForce` extraction.
- `src/cleaner.py` — Pure DataFrame filtering, rounding, near-zero removal, and sub-case aggregation.
- `src/nbcc_combos.py` — Declarative NBCC 2020 ULS superposition and governing envelope summary (`Max`/`Min` `P`, `V2`, `M3` + governing combo).
- `src/config.py` — JSON preset serialization/deserialization with graceful handling of missing model items.
- `src/exporter.py` — Formatted 3-sheet Excel export (`Envelope Summary`, `Load Combinations`, `Metadata`) and clipboard copy.
- `src/ui.py` — `tkinter`/`ttk` interface, in-memory raw data caching, and background threading for ETABS pulls.