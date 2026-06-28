"""
工作簿: Global_Hg_Flow/2.5 IND_NFMO_AlS.xlsx
工作表: 'Sum_AlWaste'
已用区域: A1:I232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Wastes from Al Production Chain |  |  |  |  |  |  |  |  | 

实现说明: C 列 VLOOKUP 外部簿 [2]Emission-r；G–I 引用 Sum_AlS 对应汞流列。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.5 IND_NFMO_AlS.xlsx", "Sum_AlWaste")
