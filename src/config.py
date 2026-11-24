"""Configuration models, JSON persistence, and model-reconciliation helpers.

Pure Python (no COM/ETABS imports) so configuration serialization and
reconciliation remain 100% testable without ETABS running.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
import math
from pathlib import Path
from typing import Any, Sequence

NBCC_CATEGORIES: tuple[str, ...] = (
    "Dead Load",
    "Live Load",
    "Snow Load",
    "Wind Load",
    "Seismic Load",
)
NBCC_CATEGORY_SYMBOLS: dict[str, str] = {
    "Dead Load": "D",
    "Live Load": "L",
    "Snow Load": "S",
    "Wind Load": "W",
    "Seismic Load": "E",
}
_CATEGORY_ALIASES: dict[str, str] = {
    "D": "Dead Load",
    "DEAD": "Dead Load",
    "DEAD LOAD": "Dead Load",
    "L": "Live Load",
    "LIVE": "Live Load",
    "LIVE LOAD": "Live Load",
    "S": "Snow Load",
    "SNOW": "Snow Load",
    "SNOW LOAD": "Snow Load",
    "W": "Wind Load",
    "WIND": "Wind Load",
    "WIND LOAD": "Wind Load",
    "E": "Seismic Load",
    "EQ": "Seismic Load",
    "SEISMIC": "Seismic Load",
    "SEISMIC LOAD": "Seismic Load",
    "EARTHQUAKE": "Seismic Load",
    "EARTHQUAKE LOAD": "Seismic Load",
}


def normalize_nbcc_category(value: str) -> str:
    """Normalize a full category name or legacy shorthand (D/L/S/W/E) to canonical form."""
    if not isinstance(value, str):
        raise ValueError(
            f"Invalid NBCC category {value!r}; expected one of {NBCC_CATEGORIES}."
        )
    key = " ".join(value.strip().upper().split())
    canonical = _CATEGORY_ALIASES.get(key)
    if canonical is None:
        raise ValueError(
            f"Invalid NBCC category {value!r}; expected one of {NBCC_CATEGORIES}."
        )
    return canonical


STATION_FILTERS: tuple[str, ...] = ("All", "Bottom", "Top")
UNIT_SYSTEMS: tuple[str, ...] = (
    "kN, m, C",
    "N, mm, C",
    "kip, ft, F",
    "kip, in, F",
)
DEFAULT_UNITS = "kN, m, C"
SCHEMA_VERSION = 1


@dataclass
class CleaningOptions:
    """User-configurable data cleaning and filtering options."""

    selected_piers: list[str] | None = None  # None => All piers
    selected_stories: list[str] | None = None  # None => All stories
    station_filter: str = "Bottom"
    round_decimals: int = 2
    drop_near_zero: bool = False
    near_zero_threshold: float = 1e-4
    remove_duplicates: bool = False
    aggregate_subcases: bool = True
    zero_lateral_axial: bool = False
    zero_gravity_shear_moment: bool = False

    def validate(self) -> None:
        if self.station_filter not in STATION_FILTERS:
            raise ValueError(
                f"Invalid station_filter {self.station_filter!r}; "
                f"expected one of {STATION_FILTERS}."
            )
        if not isinstance(self.round_decimals, int) or isinstance(
            self.round_decimals, bool
        ) or self.round_decimals < 0:
            raise ValueError(
                f"round_decimals must be a non-negative integer, got {self.round_decimals!r}."
            )
        if self.near_zero_threshold < 0:
            raise ValueError(
                f"near_zero_threshold must be non-negative, got {self.near_zero_threshold!r}."
            )


@dataclass
class ComboDefinition:
    """Editable category-level load combination definition."""

    name: str
    factors: dict[str, float] = field(default_factory=dict)
    enabled: bool = True
    case_id: str = "Custom"

    def validate(self) -> None:
        if not isinstance(self.name, str) or not self.name.strip():
            raise ValueError("Combination name must be a non-empty string.")
        if not isinstance(self.factors, dict):
            raise ValueError(f"Factors for combo {self.name!r} must be a dict.")
        cleaned: dict[str, float] = {}
        for cat, raw_val in self.factors.items():
            try:
                canonical_cat = normalize_nbcc_category(cat)
            except ValueError as exc:
                raise ValueError(
                    f"Invalid NBCC category {cat!r} in combination {self.name!r}; "
                    f"expected one of {NBCC_CATEGORIES}."
                ) from exc
            val = float(raw_val)
            if not math.isfinite(val):
                raise ValueError(
                    f"Non-finite factor {raw_val!r} for category {cat!r} in {self.name!r}."
                )
            if val != 0.0:
                cleaned[canonical_cat] = val
        if not cleaned:
            raise ValueError(
                f"Combination {self.name!r} must have at least one non-zero load factor."
            )
        self.name = self.name.strip()
        self.factors = cleaned


@dataclass
class NBCCOptions:
    """NBCC 2020 Table 4.1.3.2.-A load combination options."""

    storage_assembly_occupancy: bool = False
    custom_combinations: list[ComboDefinition] | None = None

    def validate(self) -> None:
        if self.custom_combinations is not None:
            for combo in self.custom_combinations:
                combo.validate()


@dataclass
class AppConfig:
    """Serializable project preset configuration."""

    schema_version: int = SCHEMA_VERSION
    units: str = DEFAULT_UNITS
    case_mappings: dict[str, str] = field(default_factory=dict)
    cleaning: CleaningOptions = field(default_factory=CleaningOptions)
    nbcc: NBCCOptions = field(default_factory=NBCCOptions)

    def validate(self) -> None:
        if self.units not in UNIT_SYSTEMS:
            raise ValueError(
                f"Invalid units {self.units!r}; expected one of {UNIT_SYSTEMS}."
            )
        normalized_mappings: dict[str, str] = {}
        for case_name, category in self.case_mappings.items():
            if not isinstance(case_name, str) or not case_name.strip():
                raise ValueError(f"Invalid load case name: {case_name!r}.")
            try:
                normalized_mappings[case_name] = normalize_nbcc_category(category)
            except ValueError as exc:
                raise ValueError(
                    f"Invalid NBCC category {category!r} for load case {case_name!r}; "
                    f"expected one of {NBCC_CATEGORIES}."
                ) from exc
        self.case_mappings = normalized_mappings
        self.cleaning.validate()
        self.nbcc.validate()

    def to_dict(self) -> dict[str, Any]:
        self.validate()
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> AppConfig:
        if not isinstance(data, dict):
            raise ValueError("Configuration JSON root must be an object.")

        cleaning_raw = data.get("cleaning", {})
        if not isinstance(cleaning_raw, dict):
            raise ValueError("'cleaning' section must be a JSON object.")

        nbcc_raw = data.get("nbcc", {})
        if not isinstance(nbcc_raw, dict):
            raise ValueError("'nbcc' section must be a JSON object.")

        case_mappings_raw = data.get("case_mappings", {})
        if not isinstance(case_mappings_raw, dict):
            raise ValueError("'case_mappings' must be a JSON object.")

        selected_piers = cleaning_raw.get("selected_piers")
        if selected_piers is not None:
            if not isinstance(selected_piers, list) or not all(
                isinstance(p, str) for p in selected_piers
            ):
                raise ValueError("'selected_piers' must be null or a list of strings.")
            selected_piers = list(selected_piers)

        selected_stories = cleaning_raw.get("selected_stories")
        if selected_stories is not None:
            if not isinstance(selected_stories, list) or not all(
                isinstance(s, str) for s in selected_stories
            ):
                raise ValueError("'selected_stories' must be null or a list of strings.")
            selected_stories = list(selected_stories)

        cleaning = CleaningOptions(
            selected_piers=selected_piers,
            selected_stories=selected_stories,
            station_filter=cleaning_raw.get("station_filter", "Bottom"),
            round_decimals=cleaning_raw.get("round_decimals", 2),
            drop_near_zero=bool(cleaning_raw.get("drop_near_zero", False)),
            near_zero_threshold=float(cleaning_raw.get("near_zero_threshold", 1e-4)),
            remove_duplicates=bool(cleaning_raw.get("remove_duplicates", False)),
            aggregate_subcases=bool(cleaning_raw.get("aggregate_subcases", True)),
            zero_lateral_axial=bool(cleaning_raw.get("zero_lateral_axial", False)),
            zero_gravity_shear_moment=bool(
                cleaning_raw.get("zero_gravity_shear_moment", False)
            ),
        )
        custom_combos_raw = nbcc_raw.get("custom_combinations")
        custom_combos: list[ComboDefinition] | None = None
        if custom_combos_raw is not None:
            if not isinstance(custom_combos_raw, list):
                raise ValueError("'nbcc.custom_combinations' must be null or a list.")
            custom_combos = []
            for item in custom_combos_raw:
                if not isinstance(item, dict):
                    raise ValueError(
                        "Each entry in 'nbcc.custom_combinations' must be an object."
                    )
                factors_raw = item.get("factors", {})
                if not isinstance(factors_raw, dict):
                    raise ValueError("Combination 'factors' must be a JSON object.")
                custom_combos.append(
                    ComboDefinition(
                        name=str(item.get("name", "")),
                        factors={str(k): float(v) for k, v in factors_raw.items()},
                        enabled=bool(item.get("enabled", True)),
                        case_id=str(item.get("case_id", "Custom")),
                    )
                )

        nbcc = NBCCOptions(
            storage_assembly_occupancy=bool(
                nbcc_raw.get("storage_assembly_occupancy", False)
            ),
            custom_combinations=custom_combos,
        )
        config = cls(
            schema_version=int(data.get("schema_version", SCHEMA_VERSION)),
            units=str(data.get("units", DEFAULT_UNITS)),
            case_mappings={str(k): str(v) for k, v in case_mappings_raw.items()},
            cleaning=cleaning,
            nbcc=nbcc,
        )
        config.validate()
        return config


@dataclass
class ReconciliationResult:
    """Outcome of matching a loaded config against an active ETABS model."""

    config: AppConfig
    missing_cases: list[str] = field(default_factory=list)
    missing_piers: list[str] = field(default_factory=list)
    missing_stories: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def has_warnings(self) -> bool:
        return bool(self.warnings)


def save_config(config: AppConfig, path: str | Path) -> Path:
    """Validate and serialize an AppConfig to a JSON file."""
    target = Path(path)
    payload = config.to_dict()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return target


def load_config(path: str | Path) -> AppConfig:
    """Load and validate an AppConfig from a JSON file."""
    source = Path(path)
    try:
        raw = json.loads(source.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in configuration file {source}: {exc}") from exc
    return AppConfig.from_dict(raw)


def reconcile_config_with_model(
    config: AppConfig,
    available_cases: Sequence[str],
    available_piers: Sequence[str],
    available_stories: Sequence[str],
) -> ReconciliationResult:
    """Match saved preset items against the currently connected ETABS model.

    Any load cases, piers, or stories that do not exist in the active model are
    excluded from the reconciled config and reported as non-blocking warnings.
    """
    case_set = set(available_cases)
    pier_set = set(available_piers)
    story_set = set(available_stories)

    matched_cases: dict[str, str] = {}
    missing_cases: list[str] = []
    for case_name, category in config.case_mappings.items():
        if case_name in case_set:
            matched_cases[case_name] = category
        else:
            missing_cases.append(case_name)

    missing_piers: list[str] = []
    matched_piers: list[str] | None
    if config.cleaning.selected_piers is None:
        matched_piers = None
    else:
        matched_piers = [p for p in config.cleaning.selected_piers if p in pier_set]
        missing_piers = [p for p in config.cleaning.selected_piers if p not in pier_set]

    missing_stories: list[str] = []
    matched_stories: list[str] | None
    if config.cleaning.selected_stories is None:
        matched_stories = None
    else:
        matched_stories = [
            s for s in config.cleaning.selected_stories if s in story_set
        ]
        missing_stories = [
            s for s in config.cleaning.selected_stories if s not in story_set
        ]

    warnings: list[str] = []
    if missing_cases:
        warnings.append(
            f"Load case(s) not found in active model (skipped): {', '.join(missing_cases)}"
        )

    if missing_piers:
        if matched_piers == [] and config.cleaning.selected_piers:
            matched_piers = None
            warnings.append(
                "None of the saved pier filter(s) exist in the active model "
                f"({', '.join(missing_piers)}); pier filter reset to All."
            )
        else:
            warnings.append(
                f"Pier(s) not found in active model (skipped): {', '.join(missing_piers)}"
            )

    if missing_stories:
        if matched_stories == [] and config.cleaning.selected_stories:
            matched_stories = None
            warnings.append(
                "None of the saved story filter(s) exist in the active model "
                f"({', '.join(missing_stories)}); story filter reset to All."
            )
        else:
            warnings.append(
                f"Story(s) not found in active model (skipped): {', '.join(missing_stories)}"
            )

    reconciled_cleaning = CleaningOptions(
        selected_piers=matched_piers,
        selected_stories=matched_stories,
        station_filter=config.cleaning.station_filter,
        round_decimals=config.cleaning.round_decimals,
        drop_near_zero=config.cleaning.drop_near_zero,
        near_zero_threshold=config.cleaning.near_zero_threshold,
        remove_duplicates=config.cleaning.remove_duplicates,
        aggregate_subcases=config.cleaning.aggregate_subcases,
        zero_lateral_axial=config.cleaning.zero_lateral_axial,
        zero_gravity_shear_moment=config.cleaning.zero_gravity_shear_moment,
    )
    reconciled = AppConfig(
        schema_version=config.schema_version,
        units=config.units,
        case_mappings=matched_cases,
        cleaning=reconciled_cleaning,
        nbcc=NBCCOptions(
            storage_assembly_occupancy=config.nbcc.storage_assembly_occupancy,
            custom_combinations=(
                [
                    ComboDefinition(
                        name=c.name,
                        factors=dict(c.factors),
                        enabled=c.enabled,
                        case_id=c.case_id,
                    )
                    for c in config.nbcc.custom_combinations
                ]
                if config.nbcc.custom_combinations is not None
                else None
            ),
        ),
    )

    return ReconciliationResult(
        config=reconciled,
        missing_cases=missing_cases,
        missing_piers=missing_piers,
        missing_stories=missing_stories,
        warnings=warnings,
    )
