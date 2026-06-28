"""按当前工作簿的 sheet 顺序执行某个模块目录下的脚本。

可作为模块导入（``run_folder(Path)``），也可作为 CLI 直接调用：

    python3 Code/_sheet_runner.py "Code/2.7 IND_CEM"

各模块目录无需再放入口文件，本脚本是唯一入口。
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

import openpyxl

_REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_REPO / "Code"))
from _wb_resolver import resolve_workbook_path  # noqa: E402

BOOK_RE = re.compile(r"工作簿:\s*(?:Global_Hg_Flow|Code-2022)/(.+?\.xlsx)")
SHEET_RE = re.compile(r"工作表:\s*'([^']+)'")

# 共享脚本本身不参与 sheet 调度（以下划线开头的辅助模块）
_SKIP_NAMES = {"_hg_expand.py", "_wb_resolver.py", "_sheet_runner.py"}


def parse_mapping(script: Path) -> tuple[str, str] | None:
    text = script.read_text(encoding="utf-8", errors="ignore")
    mb = BOOK_RE.search(text)
    ms = SHEET_RE.search(text)
    if not (mb and ms):
        return None
    return mb.group(1), ms.group(1)


def collect_scripts(folder: Path) -> tuple[str, Path, list[tuple[str, Path]], list[str]]:
    mapped: list[tuple[str, str, Path]] = []
    for script in sorted(folder.glob("*.py")):
        if script.name in _SKIP_NAMES:
            continue
        parsed = parse_mapping(script)
        if parsed is None:
            continue
        workbook, sheet = parsed
        mapped.append((workbook, sheet, script))

    if not mapped:
        raise RuntimeError(f"{folder} 下没有可映射到工作簿/工作表的脚本")

    workbooks = {wb for wb, _, _ in mapped}
    if len(workbooks) != 1:
        details = ", ".join(sorted(workbooks))
        raise RuntimeError(f"{folder} 下检测到多个工作簿映射: {details}")

    workbook_declared = mapped[0][0]
    wb_path = resolve_workbook_path(_REPO, workbook_declared)
    wb = openpyxl.load_workbook(wb_path, read_only=True)

    sheet_to_script = {sheet: script for _, sheet, script in mapped}
    ordered = [(sheet, sheet_to_script[sheet]) for sheet in wb.sheetnames if sheet in sheet_to_script]
    missing = sorted(set(sheet_to_script) - {sheet for sheet, _ in ordered})
    wb.close()

    return workbook_declared, wb_path, ordered, missing


def run_folder(folder: Path) -> int:
    """执行 ``folder`` 下的所有脚本，按当前工作簿 sheet 顺序。"""
    workbook_declared, wb_path, ordered, missing = collect_scripts(folder)

    print(f"[RUNNER] workbook: {workbook_declared} -> {wb_path.name}")
    if missing:
        print("[WARN] 以下脚本对应 sheet 在当前工作簿中不存在，将跳过:")
        for sh in missing:
            print(f"       - {sh}")

    for idx, (sheet, script) in enumerate(ordered, 1):
        print(f"[{idx:02d}/{len(ordered):02d}] {sheet} -> {script.name}")
        result = subprocess.run([sys.executable, str(script)], cwd=_REPO)
        if result.returncode != 0:
            print(f"[FAILED] {script.name} exit={result.returncode}")
            return result.returncode

    print("[DONE] all sheets executed")
    return 0


def _cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按当前工作簿的 sheet 顺序执行指定模块目录下的脚本。",
    )
    parser.add_argument(
        "folder",
        help="模块目录路径，例如 'Code/2.7 IND_CEM'（相对路径基于仓库根目录解析）",
    )
    args = parser.parse_args(argv)

    folder = Path(args.folder).expanduser()
    if not folder.is_absolute():
        folder = (_REPO / folder).resolve()
    if not folder.is_dir():
        print(f"[ERROR] not a directory: {folder}", file=sys.stderr)
        return 2
    return run_folder(folder)


if __name__ == "__main__":
    raise SystemExit(_cli())
