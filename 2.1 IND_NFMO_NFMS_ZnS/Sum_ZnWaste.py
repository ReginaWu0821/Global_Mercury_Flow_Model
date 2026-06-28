"""
工作簿: Global_Hg_Flow/2.1 IND_NFMO_NFMS_ZnS.xlsx
工作表: 'Sum_ZnWaste'
已用区域: A1:N232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Wastes from Zn smelting chain (t) |  |  |  |  |  |  |  |  | 

实现说明: A–F 为输入；G 引用 Sum_NFMO_Zn；H–N 引用 Sum_ZnS 对应列。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.1 IND_NFMO_NFMS_ZnS.xlsx", "Sum_ZnWaste")
