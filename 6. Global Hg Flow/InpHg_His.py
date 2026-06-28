"""
工作簿: Global_Hg_Flow/6. Global Hg Flow.xlsx
工作表: 'InpHg_His'
已用区域: A1:K232
合并单元格块数: —
首行预览: Country or Aera | Abbrev | … | 历史遗留汞输入

实现说明: 第 5 行为模板行，向下展开至第 232 行。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "6. Global Hg Flow.xlsx", "InpHg_His")
