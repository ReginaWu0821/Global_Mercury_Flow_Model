"""
工作簿: Global_Hg_Flow/2.6 IND_IOM_CW_ISP.xlsx
工作表: 'Inp_Activity'
已用区域: A1:P232
合并单元格块数: —
首行预览: Country or Aera | … | 活动量与中间计算列

实现说明: H–K、P 为公式；第 3 行 H/I/J/K 均为列合计 ``=SUM(<col>4:<col>232)``。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.6 IND_IOM_CW_ISP.xlsx", "Inp_Activity")
