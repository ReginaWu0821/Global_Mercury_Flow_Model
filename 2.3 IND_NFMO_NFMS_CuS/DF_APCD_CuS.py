"""
工作簿: Global_Hg_Flow/2.3 IND_NFMO_NFMS_CuS.xlsx
工作表: 'DF_APCD_CuS'
已用区域: A1:U65
合并单元格块数: 5
首行预览: Abbrev | APCD | Distribution Factor |  |  |  |  |  |  |  |  | Application Proportion (%) | Distribution Factor for Calculation |  |  | 

实现说明: A–K 为输入；L–S = C–J × K 列应用比例/100（数据行至第 65 行；无 Pb 版 T 列求和）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.3 IND_NFMO_NFMS_CuS.xlsx", "DF_APCD_CuS")
