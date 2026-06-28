"""
工作簿: Global_Hg_Flow/1.5 COMOilGasBio_OGP_OC_NGC.xlsx
工作表: 'Sum_OilWaste'
已用区域: A1:H232
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Wastes from Oil Combustion Chain (t) |  |  |  |  |  |  |  |  | 

实现说明: A–F 为输入；G、H 引用 Sum_OilPre_OilCom 中对应油渣汞流列。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.5 COMOilGasBio_OGP_OC_NGC.xlsx", "Sum_OilWaste")
