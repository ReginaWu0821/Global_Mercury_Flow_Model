"""
工作簿: Global_Hg_Flow/3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx
工作表: 'Sum_HgAdd_Pro_Waste'
已用区域: A1:G232
合并单元格块数: 5
首行预览: Country or Aera | Abbrev | Abbrev | Wastes from HgAdd Pro (t) |  | Wastes from HgAdd Use (t) |  |  |  |  |  |  |  |  |  | 

实现说明: A–C 为输入；D/E 引用 Sum_HgAdd_Pro 的合计列；F/G 引用 Sum_HgAdd_Use。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx", "Sum_HgAdd_Pro_Waste")
