"""
工作簿: Global_Hg_Flow/2.2 IND_NFMO_NFMS_PbS.xlsx
工作表: 'DF_APCD_PbS'
已用区域: A1:U60
合并单元格块数: 5
首行预览: Abbrev | APCD | Distribution Factor |  |  |  |  |  |  |  |  | Application Proportion (%) | Distribution Factor for Calculation |  |  | 

实现说明: A–K 为输入；L–S = C–J × K 列应用比例/100；T = SUM(L:S)。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.2 IND_NFMO_NFMS_PbS.xlsx", "DF_APCD_PbS")
