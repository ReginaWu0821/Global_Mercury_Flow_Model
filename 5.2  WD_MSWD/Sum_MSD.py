"""
工作簿: Global_Hg_Flow/5.2  WD_MSWD.xlsx
工作表: 'Sum_MSWD'
已用区域: A1:R232
合并单元格块数: 4
首行预览: Municipal Solid Waste … | … | MSD 汞流量

实现说明: 第 5 行为模板行，向下展开至第 232 行；第 3 行含横向勾稽（如 M3=B3+E3+H3），快照还原。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.2  WD_MSWD.xlsx", "Sum_MSWD")
