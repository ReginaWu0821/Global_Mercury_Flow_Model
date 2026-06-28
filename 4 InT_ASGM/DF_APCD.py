"""
工作簿: Global_Hg_Flow/4 InT_ASGM.xlsx
工作表: 'DF_APCD'
已用区域: A1:S232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev | Retort Method Share (%) | Distribution Factor_concentrate_with ... |  |  |  |  | Distribution Factor_ore_no retort |  |  |  |  | Distribution Factor_concentrate_no re... | 

实现说明: D 列引用外部簿 [1]Summary-ASGM-final；E–S 为分配因子（输入）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "4 InT_ASGM.xlsx", "DF_APCD")
