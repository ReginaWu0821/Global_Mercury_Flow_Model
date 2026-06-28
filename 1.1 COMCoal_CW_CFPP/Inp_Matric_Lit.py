"""
工作簿: Global_Hg_Flow/1.1 COMCoal_CW_CFPP.xlsx
工作表: 'Inp_Matric_Lit'
已用区域: A1:GW204
合并单元格块数: 4
首行预览:  | HgCon_Coal_Loc (mg/kg) | HgCon_Coal_Imp (mg/kg) | TransMat_Coal (kt) | Pro_Coal_2017 (kt) |  |  |  |  |  |  |  |  |  |  | 

实现说明: 黄色底为输入；在此实现该 sheet 的公式/计算逻辑，可复用 openpyxl / pandas 读源表。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.1 COMCoal_CW_CFPP.xlsx", "Inp_Matric_Lit")
