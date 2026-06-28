"""
工作簿: Global_Hg_Flow/5.1  WD_IWD.xlsx
工作表: 'IWD_Rec'
已用区域: A1:I232
合并单元格块数: 3
首行预览: Country or Aera | Abbrev | Hg Recovery Process | …

实现说明: 第 5 行为模板行，向下展开至第 232 行。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.1  WD_IWD.xlsx", "IWD_Rec")
