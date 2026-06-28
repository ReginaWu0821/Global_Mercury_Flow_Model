"""
工作簿: Global_Hg_Flow/1.5 COMOilGasBio_OGP_OC_NGC.xlsx
工作表: 'Sum_GasWaste'
已用区域: A1:F232
合并单元格块数: 5
首行预览: Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Wastes from Gas Combustion Chain (t) |  |  |  |  |  |  |  |  |  | 

实现说明: A–E 为输入；F 列引用 Sum_GasPre_GasCom 气燃烧汞流。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.5 COMOilGasBio_OGP_OC_NGC.xlsx", "Sum_GasWaste")
