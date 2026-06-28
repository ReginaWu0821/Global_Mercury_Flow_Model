"""
工作簿: Global_Hg_Flow/2.6 IND_IOM_CW_ISP.xlsx
工作表: 'Sum_ISP_Waste'
已用区域: A1:O232
合并单元格块数: —
首行预览: Country or Aera | … | 固废/回路引用列

实现说明: D–J、L–N 为公式；引用 Sum_IOM_CW、Sum_ISP。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.6 IND_IOM_CW_ISP.xlsx", "Sum_ISP_Waste")
