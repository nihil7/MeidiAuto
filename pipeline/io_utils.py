from __future__ import annotations

import glob
import os
from pathlib import Path


def resolve_data_dir(argv_value: str | None, default_dir: str = "data") -> Path:
    if argv_value:
        return Path(argv_value).resolve()
    return Path(os.getcwd(), default_dir).resolve()


def find_first_excel(data_dir: Path, pattern: str) -> str | None:
    files = glob.glob(str(data_dir / pattern))
    return files[0] if files else None


def find_latest_excel(data_dir: Path, pattern: str) -> str | None:
    files = glob.glob(str(data_dir / pattern))
    return max(files, key=os.path.getmtime) if files else None


def ensure_existing_dir(path: Path, label: str = "目录") -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label}不存在: {path}")
    return path


def find_required_excel(data_dir: Path, pattern: str, latest: bool = False) -> str:
    finder = find_latest_excel if latest else find_first_excel
    found = finder(data_dir, pattern)
    if not found:
        raise FileNotFoundError(f"没有找到符合条件的文件: {pattern}")
    return found


def find_first_match(search_dirs: list[Path], pattern: str) -> str | None:
    for search_dir in search_dirs:
        files = glob.glob(str(search_dir / pattern))
        if files:
            return files[0]
    return None


def find_required_file(search_dirs: list[Path], pattern: str, label: str = "文件") -> str:
    found = find_first_match(search_dirs, pattern)
    if not found:
        dirs_text = ", ".join(str(path) for path in search_dirs)
        raise FileNotFoundError(f"没有找到{label}: {pattern}，搜索目录: {dirs_text}")
    return found
