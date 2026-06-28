"""
工作簿: Global_Hg_Flow/2.6 IND_IOM_CW_ISP.xlsx
工作表: 'Inp_ISP'
已用区域: A1:H233
合并单元格块数: —
首行预览: Country or Aera | … | 汞输入路径列

实现说明: C–H 为公式；H 列引用外部簿 [1]Sum_COKE。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.6 IND_IOM_CW_ISP.xlsx", "Inp_ISP")
