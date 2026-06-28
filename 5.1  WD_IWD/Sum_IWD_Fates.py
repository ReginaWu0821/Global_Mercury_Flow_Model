"""
工作簿: Global_Hg_Flow/5.1  WD_IWD.xlsx
工作表: 'Sum_IWD_Fates'
已用区域: A1:IQ232
合并单元格块数: 40
首行预览: Country or Aera | Abbrev | … | 工业固废汞去向

实现说明: 第 4 行使用独立模板（如 COLUMN(J2)）；第 5–232 行由第 5 行模板展开；第 3 行为快照还原。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.1  WD_IWD.xlsx", "Sum_IWD_Fates")
