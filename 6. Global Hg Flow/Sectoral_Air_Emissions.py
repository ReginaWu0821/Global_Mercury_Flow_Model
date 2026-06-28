"""
工作簿: Global_Hg_Flow/6. Global Hg Flow.xlsx
工作表: 'Sectoral Air Emissions'
已用区域: （当前工作簿内该表无有效占用格，仅为占位）

实现说明: 源簿中该工作表暂无公式与数据；脚本打开并保存工作簿以保持与其它模块一致的入口。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "6. Global Hg Flow.xlsx", "Sectoral Air Emissions")
