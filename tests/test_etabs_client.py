"""Unit tests for src.etabs_client using a lightweight fake SapModel."""

from __future__ import annotations

import pytest

from src.etabs_client import (
    ETABSConnectionError,
    ETABSResultsError,
    UNIT_ENUM_MAP,
    _extract_model_info_from_sap_model,
    _extract_pier_forces_from_sap_model,
)


class _FakeNameListAPI:
    def __init__(self, names: list[str], ret: int = 0) -> None:
        self._names = names
        self._ret = ret

    def GetNameList(self):
        return (len(self._names), tuple(self._names), self._ret)


class _FakeLoadCasesAPI(_FakeNameListAPI):
    def GetTypeOAPI_1(self, case_name: str):
        code = 4 if case_name.startswith("RS") else 1
        return (code, 0, 1, 0)


class _FakeSetupAPI:
    def __init__(self) -> None:
        self.deselected = False
        self.selected: list[str] = []
        self.multi_step_option: int | None = None

    def DeselectAllCasesAndCombosForOutput(self) -> int:
        self.deselected = True
        self.selected.clear()
        return 0

    def SetCaseSelectedForOutput(self, name: str) -> int:
        self.selected.append(name)
        return 0

    def SetOptionMultiStepStatic(self, opt: int) -> int:
        self.multi_step_option = opt
        return 0


class _FakeResultsAPI:
    def __init__(self, rows: list[tuple], ret: int = 0) -> None:
        self.Setup = _FakeSetupAPI()
        self._rows = rows
        self._ret = ret

    def PierForce(self):
        selected = set(self.Setup.selected)
        filtered = [r for r in self._rows if r[2] in selected]
        if not filtered:
            return (0, (), (), (), (), (), (), (), (), (), (), self._ret)
        cols = list(zip(*filtered))
        return (len(filtered), *cols, self._ret)


class _FakeSapModel:
    def __init__(
        self,
        filename: str = "DemoModel.EDB",
        filepath: str = r"C:\Models\DemoModel.EDB",
        rows: list[tuple] | None = None,
        pier_force_ret: int = 0,
    ) -> None:
        self._filename = filename
        self._filepath = filepath
        self.units_set: int | None = None
        self.LoadCases = _FakeLoadCasesAPI(["Dead", "Live", "RSX"])
        self.PierLabel = _FakeNameListAPI(["P1", "P2"])
        self.Story = _FakeNameListAPI(["L2", "Roof"])
        self.Results = _FakeResultsAPI(
            rows
            if rows is not None
            else [
                ("Roof", "P1", "Dead", "Top", -100.0, 1.0, 0.5, 0.2, 3.0, 4.0),
                ("Roof", "P1", "Dead", "Bottom", -110.0, 1.0, 0.5, 0.2, 3.5, 5.0),
                ("Roof", "P1", "RSX", "Bottom", 40.0, 200.0, 10.0, 15.0, 30.0, 600.0),
            ],
            ret=pier_force_ret,
        )

    def GetModelFilename(self, include_path: bool = True) -> str:
        return self._filepath if include_path else self._filename

    def SetPresentUnits(self, unit_enum: int) -> int:
        self.units_set = unit_enum
        return 0


def test_extract_model_info_queries_metadata_and_sets_units() -> None:
    sap = _FakeSapModel()
    info = _extract_model_info_from_sap_model(sap, units="kN, m, C")

    assert sap.units_set == UNIT_ENUM_MAP["kN, m, C"]
    assert info.filename == "DemoModel.EDB"
    assert info.load_cases == ["Dead", "Live", "RSX"]
    assert info.pier_labels == ["P1", "P2"]
    assert info.stories == ["L2", "Roof"]
    assert info.case_types["RSX"] == "ResponseSpectrum"


def test_extract_model_info_raises_when_no_model_open() -> None:
    sap = _FakeSapModel(filename="", filepath="")
    with pytest.raises(ETABSConnectionError, match="no model file"):
        _extract_model_info_from_sap_model(sap)


def test_extract_pier_forces_selects_only_requested_cases() -> None:
    sap = _FakeSapModel()
    info, df = _extract_pier_forces_from_sap_model(
        sap, selected_cases=["Dead", "RSX"], units="kip, ft, F"
    )

    assert sap.units_set == UNIT_ENUM_MAP["kip, ft, F"]
    assert sap.Results.Setup.deselected is True
    assert sap.Results.Setup.selected == ["Dead", "RSX"]
    assert sap.Results.Setup.multi_step_option == 2
    assert len(df) == 3
    assert set(df["LoadCase"]) == {"Dead", "RSX"}
    assert info.units == "kip, ft, F"


def test_extract_pier_forces_raises_when_unanalyzed_or_empty() -> None:
    sap_empty = _FakeSapModel(rows=[])
    with pytest.raises(ETABSResultsError, match="0 pier force rows"):
        _extract_pier_forces_from_sap_model(sap_empty, selected_cases=["Dead"])

    sap_err = _FakeSapModel(pier_force_ret=1)
    with pytest.raises(ETABSResultsError, match="error code 1"):
        _extract_pier_forces_from_sap_model(sap_err, selected_cases=["Dead"])
