"""
工作簿: Global_Hg_Flow/3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx
工作表: 'Hg_Inpt'
已用区域: A1:M232
合并单元格块数: 2
首行预览:  | Hg Input_Production (t) |  |  |  |  |  | Hg Input_Use_considering lifecycle (t) |  |  |  |  |  |  |  | 

实现说明: B–G 列为各产品汞输入量（原始/外部链接输入）；第 3 行为合计；H–M 引用外部簿 [1]活动水平-2050。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx", "Hg_Inpt")
