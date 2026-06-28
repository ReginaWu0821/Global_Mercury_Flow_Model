"""
工作簿: Global_Hg_Flow/2.4 IND_NFMO_NFMS_LSGP.xlsx
工作表: 'DF_LSGP'
已用区域: A1:S11
合并单元格块数: 6
首行预览: Technology | Abbrev | APCD | Distribution factor |  |  |  |  |  |  | Application Proportion (%) | Distribution Factor for calculation |  |  |  | 

实现说明: A–D、技术列与 E–J 分配因子、K 应用比例为输入；M–R = E–J × K/100（第 3–11 行）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.4 IND_NFMO_NFMS_LSGP.xlsx", "DF_LSGP")
