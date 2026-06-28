"""
工作簿: Global_Hg_Flow/5.1  WD_IWD.xlsx
工作表: 'Sum_Secondary Emission'
已用区域: A1:BT232
合并单元格块数: 12
首行预览: Country or Aera | Abbrev | … | 次生排放

实现说明: 第 5 行为模板行，向下展开至第 232 行；工作表名含空格。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.1  WD_IWD.xlsx", "Sum_Secondary Emission")
