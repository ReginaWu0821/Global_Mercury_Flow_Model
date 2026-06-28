"""
工作簿: Global_Hg_Flow/4 InT_ASGM.xlsx
工作表: 'Inp_Hg Input'
已用区域: A1:F232
合并单元格块数: 4
首行预览: Country or Aera | Abbrev | Abbrev | Hg Input (t) |  |  |  |  |  |  |  |  |  |  |  | 

实现说明: A–C 为输入；D–F 由 Inp_Hg Content 与 Inp_Activities 计算；第 3 行为 D–F 合计。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "4 InT_ASGM.xlsx", "Inp_Hg Input")
