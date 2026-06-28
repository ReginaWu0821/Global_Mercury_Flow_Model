"""
工作簿: Global_Hg_Flow/3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx
工作表: 'Sum_HgAdd_Use'
已用区域: A1:F232
合并单元格块数: 1
首行预览:  | HgAdd_Use_Fate (t) |  |  |  |  |  |  |  |  |  |  |  |  |  | 

实现说明: B–F 为 Hg_Inpt 使用环节汞输入合计 × DF_APCD_Use 对应列；第 3 行为列合计。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx", "Sum_HgAdd_Use")
