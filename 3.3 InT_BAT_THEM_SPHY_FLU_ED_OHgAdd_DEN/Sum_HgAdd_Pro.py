"""
工作簿: Global_Hg_Flow/3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx
工作表: 'Sum_HgAdd_Pro'
已用区域: A1:AJ232
合并单元格块数: 7
首行预览:  | Battery |  |  |  |  | Thermometer/Sphygmomanometer |  |  |  |  | Fluorescent Lamp |  |  |  | 

实现说明: B–AE 为各产品汞输入 × DF_APCD_Pro 对应行；AF–AJ 为去向合计；第 3 行为各列 SUM。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx", "Sum_HgAdd_Pro")
