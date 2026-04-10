from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class PipelineStep:
    """A script step executed by the orchestrator."""

    filename: str
    required: bool = True
    input_patterns: tuple[str, ...] = field(default_factory=tuple)
    output_patterns: tuple[str, ...] = field(default_factory=tuple)
    support_patterns: tuple[str, ...] = field(default_factory=tuple)
