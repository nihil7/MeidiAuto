from __future__ import annotations

import argparse
import csv
import io
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CATALOG_JSON = ROOT / "docs" / "module_catalog.json"
CATALOG_MD = ROOT / "docs" / "MODULE_CATALOG.md"
CATALOG_CSV = ROOT / "docs" / "module_catalog.csv"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成/校验模块维护表（Markdown + CSV）")
    parser.add_argument("--check", action="store_true", help="仅校验产物是否最新，不写文件")
    return parser.parse_args()


def load_catalog() -> dict:
    return json.loads(CATALOG_JSON.read_text(encoding="utf-8"))


def render_markdown(data: dict) -> str:
    active = data["active_modules"]
    legacy = data["legacy_or_tooling_modules"]

    lines: list[str] = []
    lines.append("# 模块清单（自动生成）")
    lines.append("")
    lines.append("本文件由 `tools/generate_module_catalog.py` 自动生成，请勿手工修改。")
    lines.append("")
    lines.append("## 主流程模块")
    lines.append("")
    lines.append("| Step | 文件 | 分类 | 功能 | 输入 | 输出 |")
    lines.append("|---|---|---|---|---|---|")

    for row in active:
        lines.append(
            "| {step} | `{file}` | {category} | {purpose} | {inputs} | {outputs} |".format(**row)
        )

    lines.append("")
    lines.append("## 历史/工具模块")
    lines.append("")
    lines.append("| 文件 | 状态 | 说明 |")
    lines.append("|---|---|---|")
    for row in legacy:
        lines.append("| `{file}` | {status} | {reason} |".format(**row))
    lines.append("")
    return "\n".join(lines)


def render_csv_rows(data: dict) -> list[list[str]]:
    rows: list[list[str]] = [["scope", "step", "file", "category_or_status", "purpose_or_reason", "inputs", "outputs"]]

    for row in data["active_modules"]:
        rows.append([
            "active",
            row["step"],
            row["file"],
            row["category"],
            row["purpose"],
            row["inputs"],
            row["outputs"],
        ])

    for row in data["legacy_or_tooling_modules"]:
        rows.append([
            "legacy_or_tooling",
            "",
            row["file"],
            row["status"],
            row["reason"],
            "",
            "",
        ])

    return rows


def _csv_text(rows: list[list[str]]) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    writer.writerows(rows)
    return buf.getvalue()


def write_csv(path: Path, rows: list[list[str]]) -> None:
    path.write_text(_csv_text(rows), encoding="utf-8")


def main() -> int:
    args = parse_args()
    data = load_catalog()

    md = render_markdown(data)
    rows = render_csv_rows(data)

    current_md = CATALOG_MD.read_text(encoding="utf-8") if CATALOG_MD.exists() else ""

    current_csv = ""
    if CATALOG_CSV.exists():
        current_csv = CATALOG_CSV.read_text(encoding="utf-8")

    expected_csv = _csv_text(rows)

    if args.check:
        ok = (md == current_md) and (expected_csv == current_csv)
        if ok:
            print("✅ 模块表产物已是最新")
            return 0
        print("❌ 模块表产物不是最新，请执行: python tools/generate_module_catalog.py")
        return 1

    CATALOG_MD.write_text(md, encoding="utf-8")
    write_csv(CATALOG_CSV, rows)

    print(f"✅ 已生成: {CATALOG_MD}")
    print(f"✅ 已生成: {CATALOG_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
