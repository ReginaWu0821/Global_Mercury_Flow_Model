"""
工作簿: Global_Hg_Flow/3.2 InT_CAP_HgCat_VCM.xlsx
工作表: 'Sum_VCM'
已用区域: A1:H232
合并单元格块数: 3
首行预览:  | Hg Input (t) | Distribution Factor |  |  | Emissin and Releases |  |  |  |  |  |  |  |  |  | 

实现说明: B–E 为输入；F–H 为 B 与 C–E 乘积。B3 为手工汇总量，脚本不覆盖。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.2 InT_CAP_HgCat_VCM.xlsx", "Sum_VCM")
