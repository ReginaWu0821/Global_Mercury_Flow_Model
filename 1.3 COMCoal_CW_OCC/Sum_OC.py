"""
工作簿: Global_Hg_Flow/1.3 COMCoal_CW_OCC.xlsx
工作表: 'Sum_OC'
已用区域: A1:AG232
合并单元格块数: —
首行预览: Country or Aera | Abbrev | … | Hg 流平衡列

实现说明: 黄色底为输入；在此实现该 sheet 的公式/计算逻辑，可复用 openpyxl / pandas 读源表。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.3 COMCoal_CW_OCC.xlsx", "Sum_OC")
