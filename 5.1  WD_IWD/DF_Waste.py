"""
工作簿: Global_Hg_Flow/5.1  WD_IWD.xlsx
工作表: 'DF_Waste'
已用区域: A1:AJ142
合并单元格块数: 45
首行预览: Waste type | Country or Aera | … | 分配因子

实现说明: 第 5 行模板展开至第 4–9 行；第 10 行模板展开至第 10–30 行；第 31 行起公式形态多变，
由同目录 ``df_waste_formulas_r31.json`` 固化；第 3 行快照还原。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "5.1  WD_IWD.xlsx", "DF_Waste")
