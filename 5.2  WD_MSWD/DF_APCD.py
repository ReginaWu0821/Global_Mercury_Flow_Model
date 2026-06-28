"""
工作簿: Global_Hg_Flow/5.2  WD_MSWD.xlsx
工作表: 'DF_APCD'
已用区域: A1:V232
合并单元格块数: 8
首行预览: Municipal Waste Disposal Method | … | 分配因子

实现说明: 第 4 行为模板行，向下展开至第 232 行。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.2  WD_MSWD.xlsx", "DF_APCD")
