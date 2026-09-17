"""Validation and normalization of reimbursement run configuration."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Mapping


@dataclass(frozen=True)
class RunConfig:
    month: str
    root: Path
    output_dir: Path
    variance_limit: float = 10.0
    dry_run: bool = False

    def __post_init__(self) -> None:
        if len(self.month) != 7 or self.month[4] != "-":
            raise ValueError("month must use YYYY-MM format")
        year, month = self.month.split("-")
        if not (year.isdigit() and month.isdigit() and 1 <= int(month) <= 12):
            raise ValueError("month must use a valid YYYY-MM value")
        if self.variance_limit < 0:
            raise ValueError("variance_limit must be non-negative")


def build_config(values: Mapping[str, object]) -> RunConfig:
    """Build a validated RunConfig from CLI/config-like values."""
    root = Path(str(values.get("root", "."))).expanduser()
    output = Path(str(values.get("output_dir", root / "output"))).expanduser()
    return RunConfig(
        month=str(values.get("month", "")).strip(),
        root=root,
        output_dir=output,
        variance_limit=float(values.get("variance_limit", 10.0)),
        dry_run=bool(values.get("dry_run", False)),
    )


def config_to_dict(config: RunConfig) -> dict[str, object]:
    """Return a stable primitive representation for diagnostics."""
    return {
        "month": config.month,
        "root": str(config.root),
        "output_dir": str(config.output_dir),
        "variance_limit": config.variance_limit,
        "dry_run": config.dry_run,
    }


def resolve_output(config: RunConfig) -> Path:
    """Resolve the month-specific output directory."""
    return config.output_dir / config.month


def validate_required_paths(config: RunConfig, required: tuple[str, ...]) -> tuple[Path, ...]:
    """Validate required relative paths under the configured root."""
    resolved = []
    for relative in required:
        path = (config.root / relative).resolve()
        if not path.exists():
            raise FileNotFoundError(path)
        resolved.append(path)
    return tuple(resolved)


def ensure_output_directory(config: RunConfig) -> Path:
    """Create the configured output directory when not running dry."""
    target = resolve_output(config)
    if not config.dry_run:
        target.mkdir(parents=True, exist_ok=True)
    return target


def is_dry_run(config: RunConfig) -> bool:
    return config.dry_run
