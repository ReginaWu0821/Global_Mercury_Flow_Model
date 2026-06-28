"""
工作簿: Global_Hg_Flow/5.2  WD_MSWD.xlsx
工作表: 'Hg_Inpt'
已用区域: A1:M232
合并单元格块数: —
首行预览: Country or Aera | Abbrev | … | 外部簿链接与 MSD 输入

实现说明: 黄色底为输入；第 5 行为公式模板，向下展开；第 3 行快照还原。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.2  WD_MSWD.xlsx", "Hg_Inpt")
