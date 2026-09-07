from __future__ import annotations

import glob
from pathlib import Path
from typing import Iterable

from .models import PipelineStep


def missing_step_files(script_dir: Path, steps: Iterable[PipelineStep]) -> list[str]:
    missing: list[str] = []
    for step in steps:
        if step.required and not (script_dir / step.filename).exists():
            missing.append(step.filename)
    return missing


def _find_matches(data_dir: Path, pattern: str) -> list[str]:
    return glob.glob(str(data_dir / pattern))


def validate_step_inputs(step: PipelineStep, data_dir: Path) -> tuple[bool, str]:
    for pattern in step.input_patterns:
        if not _find_matches(data_dir, pattern):
            return False, f"缺少输入产物：{pattern}"
    return True, ""


def validate_step_output(step: PipelineStep, data_dir: Path) -> tuple[bool, str]:
    """关键步骤产物校验，防止子脚本误返回 0 导致级联失败。"""
    for pattern in step.output_patterns:
        if not _find_matches(data_dir, pattern):
            return False, f"缺少输出产物：{pattern}"
    return True, ""
