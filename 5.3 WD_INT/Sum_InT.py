"""
工作簿: Global_Hg_Flow/5.3 WD_INT.xlsx
工作表: 'Sum_InT'
已用区域: A1:E232
合并单元格块数: 3
首行预览: Country or Aera | Abbrev | Hg Fate from INT | …

实现说明: 第 5 行为模板行，向下展开至第 232 行；引用 ``Inp_Hg`` 与 ``DF_APCD``。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.3 WD_INT.xlsx", "Sum_InT")
