"""
工作簿: Global_Hg_Flow/1.5 COMOilGasBio_OGP_OC_NGC.xlsx
工作表: 'Sum_BioCom'
已用区域: A1:H232
合并单元格块数: 6
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Biomass (t) | Biomass Combustion (t) |  |  |  |  |  |  |  | 

实现说明: A–F、G 为输入；G 列汞输入来自 Inp_Hg 生物质列 O；H 为与 DF_APCD 生物质分配因子之积。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.5 COMOilGasBio_OGP_OC_NGC.xlsx", "Sum_BioCom")
