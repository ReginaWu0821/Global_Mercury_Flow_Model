"""
工作簿: Global_Hg_Flow/2.3 IND_NFMO_NFMS_CuS.xlsx
工作表: 'Sum_CuS'
已用区域: A1:O232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Emission and distribution (t) |  |  |  |  |  |  |  |  | 

实现说明: A–F、C 列为输入；G–N 为 Inp_CuS × SUMIFS(DF_APCD_CuS L–S)。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.3 IND_NFMO_NFMS_CuS.xlsx", "Sum_CuS")
