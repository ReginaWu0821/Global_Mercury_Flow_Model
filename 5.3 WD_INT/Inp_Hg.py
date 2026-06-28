"""
工作簿: Global_Hg_Flow/5.3 WD_INT.xlsx
工作表: 'Inp_Hg'
已用区域: A1:D232
合并单元格块数: —
首行预览: Country or Aera | Abbrev | … | Hg 输入（灰底为手工输入）

实现说明: 数据行为输入值；仅第 3 行 D 列为 ``=SUM(D4:D232)``。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.3 WD_INT.xlsx", "Inp_Hg")
