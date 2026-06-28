"""
工作簿: Global_Hg_Flow/3.1 InT_PHgM.xlsx
工作表: 'Sum_PHgM_Waste'
已用区域: A1:G232
合并单元格块数: 4
首行预览: Country or Aera | Abbrev | Abbrev | Wastes from PHgM Chain (t) |  |  |  |  |  |  |  |  |  |  |  | 

实现说明: A–C 为输入；D/E 引用 Sum_PHgM 中含汞渣对应列（与 Sum_PHgM 中含公式行一致）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.1 InT_PHgM.xlsx", "Sum_PHgM_Waste")
