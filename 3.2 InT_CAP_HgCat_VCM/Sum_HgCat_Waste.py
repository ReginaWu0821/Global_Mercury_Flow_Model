"""
工作簿: Global_Hg_Flow/3.2 InT_CAP_HgCat_VCM.xlsx
工作表: 'Sum_HgCat_Waste'
已用区域: A1:F232
合并单元格块数: 3
首行预览: Country or Aera | Abbrev | Abbrev | Wastes from HgCat Chain (t) |  |  |  |  |  |  |  |  |  |  |  | 

实现说明: A–C 为输入；D 引用 Sum_HgCat 列 G（Wastes_Recycle 路径）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.2 InT_CAP_HgCat_VCM.xlsx", "Sum_HgCat_Waste")
