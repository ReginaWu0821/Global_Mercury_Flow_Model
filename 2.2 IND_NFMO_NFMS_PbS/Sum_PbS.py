"""
工作簿: Global_Hg_Flow/2.2 IND_NFMO_NFMS_PbS.xlsx
工作表: 'Sum_PbS'
已用区域: A1:O232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Emission and distribution (t) |  |  |  |  |  |  |  |  | 

实现说明: A–F、C 列为输入；G–N 为 Inp_PbS × SUMIFS(DF_APCD_PbS L–S)。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.2 IND_NFMO_NFMS_PbS.xlsx", "Sum_PbS")
