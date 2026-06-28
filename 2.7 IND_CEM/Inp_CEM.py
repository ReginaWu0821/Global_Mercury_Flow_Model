"""
工作簿: Global_Hg_Flow/2.7 IND_CEM.xlsx
工作表: 'Inp_CEM'
已用区域: A1:H233
合并单元格块数: —
首行预览: Country or Aera | … | 各路径汞输入 C–H

实现说明: C–H 为公式；引用 Inp_HgContent、Inp_Activity。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.7 IND_CEM.xlsx", "Inp_CEM")
