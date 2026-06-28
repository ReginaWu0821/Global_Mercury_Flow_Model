"""
工作簿: Global_Hg_Flow/2.2 IND_NFMO_NFMS_PbS.xlsx
工作表: 'Inp_Matric_HgPb'
已用区域: A1:GZ205
合并单元格块数: 3
首行预览:  |  |  |  |  | TransMat (kt) | Production (kt, zinc metal) |  |  |  |  |  |  |  |  | 

实现说明: A、F–GY 为输入；B–E、GZ 为公式。Pb 版 VLOOKUP 范围至 I51，GZ 含 DF_NFMO 与 0.95。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.2 IND_NFMO_NFMS_PbS.xlsx", "Inp_Matric_HgPb")
