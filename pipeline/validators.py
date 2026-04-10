from __future__ import annotations

import glob
from pathlib import Path
from typing import Iterable

from .models import PipelineStep


def missing_step_files(script_dir: Path, steps: Iterable[PipelineStep], require_all: bool = False) -> list[str]:
    missing: list[str] = []
    for step in steps:
        if (require_all or step.required) and not (script_dir / step.filename).exists():
            missing.append(step.filename)
    return missing


def _find_matches(search_dir: Path, pattern: str) -> list[str]:
    return glob.glob(str(search_dir / pattern))


def validate_step_inputs(step: PipelineStep, data_dir: Path) -> tuple[bool, str]:
    for pattern in step.input_patterns:
        if not _find_matches(data_dir, pattern):
            return False, f"缺少输入产物: {pattern}"
    return True, ""


def validate_step_support_inputs(step: PipelineStep, search_dirs: Iterable[Path]) -> tuple[bool, str]:
    directories = tuple(search_dirs)
    for pattern in step.support_patterns:
        if any(_find_matches(search_dir, pattern) for search_dir in directories):
            continue
        dirs_text = ", ".join(str(path) for path in directories)
        return False, f"缺少支撑文件: {pattern}；已搜索: {dirs_text}"
    return True, ""


def validate_step_output(step: PipelineStep, data_dir: Path) -> tuple[bool, str]:
    for pattern in step.output_patterns:
        if not _find_matches(data_dir, pattern):
            return False, f"缺少输出产物: {pattern}"
    return True, ""
