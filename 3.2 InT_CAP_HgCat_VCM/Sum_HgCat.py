"""
工作簿: Global_Hg_Flow/3.2 InT_CAP_HgCat_VCM.xlsx
工作表: 'Sum_HgCat'
已用区域: A1:H232
合并单元格块数: 3
首行预览:  | Hg Input (t) | Distribution Factor |  |  | Emissin and Releases |  |  |  |  |  |  |  |  |  | 

实现说明: C–E 为分配因子输入；H 引用 Sum_VCM 汞输入；B=H/E；F、G 为 B 与 C、D 乘积（与表内中国/印度/俄罗斯行一致）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.2 InT_CAP_HgCat_VCM.xlsx", "Sum_HgCat")
