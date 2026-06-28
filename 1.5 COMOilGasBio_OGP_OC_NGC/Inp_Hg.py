"""
工作簿: Global_Hg_Flow/1.5 COMOilGasBio_OGP_OC_NGC.xlsx
工作表: 'Inp_Hg'
已用区域: A1:O233
合并单元格块数: 6
首行预览: Country or Aera | Abbrev | Abbrev | Hg Content |  |  |  | Activity Level (kt) |  |  |  | Hg Input (t) |  |  |  | 

实现说明: A–K 等为原始输入；L–O 列为由含量与活动量计算的汞输入；第 3 行为各列合计。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.5 COMOilGasBio_OGP_OC_NGC.xlsx", "Inp_Hg")
