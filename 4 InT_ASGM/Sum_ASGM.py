"""
工作簿: Global_Hg_Flow/4 InT_ASGM.xlsx
工作表: 'Sum_ASGM'
已用区域: A1:V232
合并单元格块数: 6
首行预览: Country or Aera | Abbrev | Hg Fates_concentrate_with retort |  |  |  |  | Hg Fates_ore_no retort |  |  |  |  | Hg Fates_concentrate_no retort |  |  | 

实现说明: C–Q 为汞去向细分；R–V 为三大类汇总；第 3 行先对各去向列 SUM，再由 R–V 汇总。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "4 InT_ASGM.xlsx", "Sum_ASGM")
