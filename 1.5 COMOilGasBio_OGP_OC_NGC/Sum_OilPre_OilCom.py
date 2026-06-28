"""
工作簿: Global_Hg_Flow/1.5 COMOilGasBio_OGP_OC_NGC.xlsx
工作表: 'Sum_OilPre_OilCom'
已用区域: A1:M232
合并单元格块数: 8
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Crude Oil | Crude Oil Purification (t) |  |  |  | Oil Combustion (t) |  |  |  | 

实现说明: A–F、Crude Oil(G) 等为输入或链接；H–M 为分配因子与汞输入的乘积；G 列引用 Inp_Hg 精炼汞输入。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.5 COMOilGasBio_OGP_OC_NGC.xlsx", "Sum_OilPre_OilCom")
