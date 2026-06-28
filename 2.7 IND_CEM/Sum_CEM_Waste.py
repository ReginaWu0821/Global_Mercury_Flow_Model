"""
工作簿: Global_Hg_Flow/2.7 IND_CEM.xlsx
工作表: 'Sum_CEM_Waste'
已用区域: A1:H232
合并单元格块数: —
首行预览: Country or Aera | … | 废物汞流 D–F

实现说明: D–F 引用 Sum_CEM 汇总列 O、P、R。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.7 IND_CEM.xlsx", "Sum_CEM_Waste")
