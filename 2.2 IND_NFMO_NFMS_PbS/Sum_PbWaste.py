"""
工作簿: Global_Hg_Flow/2.2 IND_NFMO_NFMS_PbS.xlsx
工作表: 'Sum_PbWaste'
已用区域: A1:N232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Wastes from Pb smelting chain (t) |  |  |  |  |  |  |  |  | 

实现说明: A–F 为输入；G 引用 Sum_NFMO_Pb!I；H–M 引用 Sum_PbS!H–M（工作簿内无 N 列废物公式）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.2 IND_NFMO_NFMS_PbS.xlsx", "Sum_PbWaste")
