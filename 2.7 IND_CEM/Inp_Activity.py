"""
工作簿: Global_Hg_Flow/2.7 IND_CEM.xlsx
工作表: 'Inp_Activity'
已用区域: A1:U232
合并单元格块数: —
首行预览: Country or Aera | … | 熟料/水泥活动量及汞输入相关列

实现说明: H–K、S、U 为公式；S 列 VLOOKUP 外部簿 [2]Hg input。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.7 IND_CEM.xlsx", "Inp_Activity")
