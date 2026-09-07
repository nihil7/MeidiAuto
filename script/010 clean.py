from __future__ import annotations

# ================================================
# STEP CARD
# 功能: 清理 data 与 script/data 下中间产物，便于重复运行。
# 输入: data 目录（argv[1] 可选）
# 输出: 清理日志（stdout）
# 上游: 主流程结束/手动触发
# 下游: 无（收尾步骤）
# ================================================

import os
import sys
from pathlib import Path

# 要删除的文件名关键字（保持历史兼容）
KEYWORDS = (
    "总库存",
    "美的仓储自动化",
    "合肥市",
    "存量查询",
    "output.html",
    "mail_meta",
    "last_mail_html",
)


def resolve_cleanup_dirs() -> list[Path]:
    """解析清理目录。

    优先清理 main.py 传入的数据目录（argv[1]），并补充历史目录：
    - 仓库 data/
    - script/data/
    """
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    candidates: list[Path] = []
    if len(sys.argv) > 1 and sys.argv[1]:
        candidates.append(Path(sys.argv[1]).expanduser().resolve())

    candidates.extend([repo_root / "data", script_dir / "data"])

    # 去重并过滤不存在目录
    unique: list[Path] = []
    seen: set[str] = set()
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        if path.exists() and path.is_dir():
            unique.append(path)
    return unique


def should_delete(filename: str) -> bool:
    return any(keyword in filename for keyword in KEYWORDS)


def clean_dir(target_dir: Path) -> int:
    deleted = 0
    for root, _, files in os.walk(target_dir):
        for filename in files:
            if not should_delete(filename):
                continue
            file_path = Path(root) / filename
            try:
                file_path.unlink()
                deleted += 1
                print(f"✅ 删除文件: {file_path}")
            except Exception as exc:
                print(f"❌ 删除文件失败: {file_path}，原因: {exc}")
    return deleted


def main() -> int:
    cleanup_dirs = resolve_cleanup_dirs()
    if not cleanup_dirs:
        print("⚠️ 未找到可清理目录（期望 data/ 或 script/data/）。")
        return 0

    print("🧹 清理目录：")
    for path in cleanup_dirs:
        print(f" - {path}")

    total_deleted = 0
    for path in cleanup_dirs:
        total_deleted += clean_dir(path)

    print(f"\n处理完成！共删除 {total_deleted} 个文件。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
