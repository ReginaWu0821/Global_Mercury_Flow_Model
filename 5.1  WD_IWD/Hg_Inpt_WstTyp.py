"""
工作簿: Global_Hg_Flow/5.1  WD_IWD.xlsx
工作表: 'Hg_Inpt_WstTyp'
已用区域: A1:BP232
合并单元格块数: 18
首行预览: Country or Aera | … | 废物类型输入

实现说明: 黄色底为输入；第 3 行为合计；第 5 行为公式模板，向下展开至第 232 行。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.1  WD_IWD.xlsx", "Hg_Inpt_WstTyp")
