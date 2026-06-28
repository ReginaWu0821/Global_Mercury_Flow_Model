"""
工作簿: Global_Hg_Flow/3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx
工作表: 'DF_APCD_Pro'
已用区域: A1:G8
合并单元格块数: 1
首行预览:  | Distribution Factor_Production |  |  |  |  |  |  |  |  |  |  |  |  |  | 

实现说明: 生产环节分配因子表；黄色区域为输入参数，表内无公式。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.3 InT_BAT_THEM_SPHY_FLU_ED_OHgAdd_DEN.xlsx", "DF_APCD_Pro")
