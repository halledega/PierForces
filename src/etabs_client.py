"""Isolated COM/OAPI client for CSI ETABS.

All `comtypes` and ETABS OAPI calls live strictly in this module. Each public
operation manages its own thread-local COM apartment (`CoInitialize` /
`CoUninitialize`) and attaches freshly to the active ETABS instance so it is
safe to invoke from Tkinter background worker threads without cross-apartment
marshaling errors.
"""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator, Sequence

import pandas as pd

from src.cleaner import REQUIRED_RAW_COLUMNS
from src.config import DEFAULT_UNITS

ETABS_PROG_ID = "CSI.ETABS.API.ETABSObject"

# Standard CSI ETABS eUnits enumeration values.
UNIT_ENUM_MAP: dict[str, int] = {
    "kip, in, F": 3,
    "kip, ft, F": 4,
    "kN, m, C": 6,
    "N, mm, C": 9,
}

# ETABS eLoadCaseType subset for informational diagnostics.
_LOAD_CASE_TYPE_NAMES: dict[int, str] = {
    1: "LinearStatic",
    2: "NonlinearStatic",
    3: "Modal",
    4: "ResponseSpectrum",
    5: "LinearHistory",
    6: "NonlinearHistory",
    7: "LinearDynamic",
    8: "NonlinearDynamic",
    9: "MovingLoad",
    10: "Buckling",
    11: "SteadyState",
    12: "PowerSpectralDensity",
    13: "LinearStaticMultiStep",
    14: "HyperStatic",
}


class ETABSError(RuntimeError):
    """Base error for ETABS connection or OAPI failures."""


class ETABSConnectionError(ETABSError):
    """Raised when no running ETABS instance or open model is found."""


class ETABSResultsError(ETABSError):
    """Raised when pier force results cannot be extracted (e.g., not analyzed)."""


@dataclass
class ETABSModelInfo:
    """Upfront metadata queried immediately upon connecting to ETABS."""

    filename: str
    filepath: str
    units: str
    load_cases: list[str] = field(default_factory=list)
    pier_labels: list[str] = field(default_factory=list)
    stories: list[str] = field(default_factory=list)
    case_types: dict[str, str] = field(default_factory=dict)


@contextmanager
def _com_apartment() -> Iterator[Any]:
    """Initialize a COM apartment for the current thread and release on exit."""
    try:
        import comtypes
        import comtypes.client
    except ImportError as exc:
        raise ETABSConnectionError(
            "comtypes is not installed in the active environment."
        ) from exc

    comtypes.CoInitialize()
    try:
        yield comtypes.client
    finally:
        comtypes.CoUninitialize()


def _attach_sap_model(com_client: Any) -> Any:
    """Attach to the active ETABS instance and return `SapModel`."""
    try:
        etabs_obj = com_client.GetActiveObject(ETABS_PROG_ID)
    except Exception as exc:
        raise ETABSConnectionError(
            "Could not attach to an open CSI ETABS instance. "
            "Please open ETABS with an active model and try again."
        ) from exc

    if etabs_obj is None:
        raise ETABSConnectionError(
            "No active CSI ETABS object was returned. Please open a model in ETABS."
        )

    sap_model = getattr(etabs_obj, "SapModel", None)
    if sap_model is None:
        raise ETABSConnectionError(
            "Connected to ETABS, but SapModel is unavailable. "
            "Please ensure a model is open."
        )
    return sap_model


def _unpack_name_list(api_obj: Any, entity_label: str) -> list[str]:
    """Call `<api_obj>.GetNameList()` safely across early/late-bound comtypes."""
    try:
        res = api_obj.GetNameList()
    except TypeError:
        res = api_obj.GetNameList(0, [])

    if not isinstance(res, (tuple, list)) or len(res) < 3:
        raise ETABSError(
            f"Unexpected return shape from {entity_label}.GetNameList(): {res!r}"
        )

    # comtypes returns (NumberNames, MyName, ret_flag).
    num_names, raw_names, ret_flag = res[0], res[1], res[-1]
    if int(ret_flag) != 0:
        raise ETABSError(
            f"ETABS failed to query {entity_label} names (error code {ret_flag})."
        )
    if not num_names or raw_names is None:
        return []
    return [str(name) for name in raw_names]


def _set_model_units(sap_model: Any, units: str) -> None:
    if units not in UNIT_ENUM_MAP:
        raise ValueError(
            f"Unsupported units {units!r}; expected one of {tuple(UNIT_ENUM_MAP)}."
        )
    ret = sap_model.SetPresentUnits(UNIT_ENUM_MAP[units])
    if int(ret) != 0:
        raise ETABSError(
            f"ETABS failed to set output units to {units!r} (error code {ret})."
        )


def _query_case_types(sap_model: Any, load_cases: Sequence[str]) -> dict[str, str]:
    """Best-effort query of ETABS load case types for diagnostics."""
    case_types: dict[str, str] = {}
    get_type = getattr(sap_model.LoadCases, "GetTypeOAPI_1", None)
    if get_type is None:
        return case_types

    for case_name in load_cases:
        try:
            res = get_type(case_name)
        except TypeError:
            try:
                res = get_type(case_name, 0, 0, 0)
            except Exception:
                continue
        except Exception:
            continue

        if isinstance(res, (tuple, list)) and len(res) >= 2:
            ret_flag = res[-1]
            if int(ret_flag) == 0:
                type_code = int(res[0])
                case_types[case_name] = _LOAD_CASE_TYPE_NAMES.get(
                    type_code, f"Type({type_code})"
                )
    return case_types


def _extract_model_info_from_sap_model(
    sap_model: Any,
    units: str = DEFAULT_UNITS,
) -> ETABSModelInfo:
    """Query active model filename, units, LoadCases, PierLabels, and Stories."""
    filepath = str(sap_model.GetModelFilename(True) or "").strip()
    filename = str(sap_model.GetModelFilename(False) or "").strip()
    if not filename and filepath:
        filename = Path(filepath).name
    if not filename:
        raise ETABSConnectionError(
            "ETABS is open, but no model file is currently loaded."
        )

    _set_model_units(sap_model, units)

    load_cases = _unpack_name_list(sap_model.LoadCases, "LoadCases")
    pier_labels = _unpack_name_list(sap_model.PierLabel, "PierLabel")
    stories = _unpack_name_list(sap_model.Story, "Story")
    case_types = _query_case_types(sap_model, load_cases)

    return ETABSModelInfo(
        filename=filename,
        filepath=filepath,
        units=units,
        load_cases=load_cases,
        pier_labels=pier_labels,
        stories=stories,
        case_types=case_types,
    )


def _extract_pier_forces_from_sap_model(
    sap_model: Any,
    selected_cases: Sequence[str],
    units: str = DEFAULT_UNITS,
) -> tuple[ETABSModelInfo, pd.DataFrame]:
    """Configure selective case output and pull PierForce results into a DataFrame."""
    if not selected_cases:
        raise ValueError("At least one load case must be selected before pulling.")

    info = _extract_model_info_from_sap_model(sap_model, units=units)

    setup = sap_model.Results.Setup
    ret = setup.DeselectAllCasesAndCombosForOutput()
    if int(ret) != 0:
        raise ETABSResultsError(
            f"ETABS failed to clear selected output cases/combos (error code {ret})."
        )

    # Request step-by-step output (2) for multi-step static cases (e.g., seismic
    # accidental eccentricity) so individual steps are preserved rather than enveloped.
    set_multi = getattr(setup, "SetOptionMultiStepStatic", None)
    if set_multi is not None:
        try:
            set_multi(2)
        except Exception:
            pass

    missing_selected = [c for c in selected_cases if c not in set(info.load_cases)]
    if missing_selected:
        raise ETABSResultsError(
            "Selected load case(s) do not exist in the active ETABS model: "
            + ", ".join(missing_selected)
        )

    for case_name in selected_cases:
        ret = setup.SetCaseSelectedForOutput(case_name)
        if int(ret) != 0:
            raise ETABSResultsError(
                f"ETABS failed to select load case {case_name!r} for output "
                f"(error code {ret})."
            )

    try:
        res = sap_model.Results.PierForce()
    except TypeError:
        res = sap_model.Results.PierForce(
            0, [], [], [], [], [], [], [], [], [], []
        )

    if not isinstance(res, (tuple, list)) or len(res) < 12:
        raise ETABSResultsError(
            f"Unexpected return shape from SapModel.Results.PierForce(): {res!r}"
        )

    (
        num_results,
        story_names,
        pier_names,
        load_cases,
        locations,
        p_vals,
        v2_vals,
        v3_vals,
        t_vals,
        m2_vals,
        m3_vals,
        ret_flag,
    ) = res[:12]

    if int(ret_flag) != 0:
        raise ETABSResultsError(
            f"ETABS PierForce extraction failed (error code {ret_flag}). "
            "Verify the model has been analyzed."
        )

    if int(num_results) == 0:
        raise ETABSResultsError(
            "ETABS returned 0 pier force rows for the selected load cases. "
            "Verify that the model has been analyzed for those cases and that "
            "wall piers are assigned."
        )

    df = pd.DataFrame(
        {
            "Story": [str(x) for x in story_names],
            "Pier": [str(x) for x in pier_names],
            "LoadCase": [str(x) for x in load_cases],
            "Location": [str(x) for x in locations],
            "P": [float(x) for x in p_vals],
            "V2": [float(x) for x in v2_vals],
            "V3": [float(x) for x in v3_vals],
            "T": [float(x) for x in t_vals],
            "M2": [float(x) for x in m2_vals],
            "M3": [float(x) for x in m3_vals],
        },
        columns=list(REQUIRED_RAW_COLUMNS),
    )

    return info, df


class ETABSClient:
    """Thread-safe façade for querying metadata and pier forces from active ETABS."""

    def connect_and_query_metadata(
        self, units: str = DEFAULT_UNITS
    ) -> ETABSModelInfo:
        """Attach to active ETABS model and return upfront metadata."""
        with _com_apartment() as com_client:
            sap_model = _attach_sap_model(com_client)
            return _extract_model_info_from_sap_model(sap_model, units=units)

    def pull_pier_forces(
        self,
        selected_cases: Sequence[str],
        units: str = DEFAULT_UNITS,
    ) -> tuple[ETABSModelInfo, pd.DataFrame]:
        """Attach to active ETABS model and pull PierForce rows for selected cases."""
        with _com_apartment() as com_client:
            sap_model = _attach_sap_model(com_client)
            return _extract_pier_forces_from_sap_model(
                sap_model,
                selected_cases=selected_cases,
                units=units,
            )
