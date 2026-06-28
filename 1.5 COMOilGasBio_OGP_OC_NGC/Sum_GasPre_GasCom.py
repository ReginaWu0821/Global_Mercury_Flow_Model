"""
工作簿: Global_Hg_Flow/1.5 COMOilGasBio_OGP_OC_NGC.xlsx
工作表: 'Sum_GasPre_GasCom'
已用区域: A1:L232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Crude Natural Gas | Crude Natural Gas Refining (t) |  |  |  | Gas Combustion (t) |  |  |  | 

实现说明: A–F、G 为输入；G 列汞输入来自 Inp_Hg；H–K 与 DF_APCD 气链对应列相乘；L 列等于 J 列（第 3 行 L3=J3，非 SUM）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.5 COMOilGasBio_OGP_OC_NGC.xlsx", "Sum_GasPre_GasCom")
