"""
工作簿: Global_Hg_Flow/6. Global Hg Flow.xlsx
工作表: 'InpHg_GeoRes'
已用区域: A1:AC232
合并单元格块数: 11
首行预览: Country or Aera | Abbrev | Reg_code | … | 燃料与矿石地理储量输入

实现说明: 黄色底为输入；第 5 行为公式模板，向下展开；第 3 行快照还原（含 AC3=AA3 等）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "6. Global Hg Flow.xlsx", "InpHg_GeoRes")
