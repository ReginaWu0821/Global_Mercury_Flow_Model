"""
工作簿: Global_Hg_Flow/2.7 IND_CEM.xlsx
工作表: 'Inp_HgContent'
已用区域: A1:K232
合并单元格块数: —
首行预览:  | Hg Content (mg/kg) | … | 折算汞含量 K

实现说明: K 列为 VLOOKUP 外部簿 [1]垃圾汞含量-折算；其余列为输入。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.7 IND_CEM.xlsx", "Inp_HgContent")
