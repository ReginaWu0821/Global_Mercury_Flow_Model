"""
工作簿: Global_Hg_Flow/2.1 IND_NFMO_NFMS_ZnS.xlsx
工作表: 'Inp_Matric_HgZn'
已用区域: A1:GZ205
合并单元格块数: 3
首行预览:  |  |  |  |  | TransMat (kt) | Production (kt, zinc metal) |  |  |  |  |  |  |  |  | 

实现说明: A、F–GY 行为原始输入；B–E、GZ 为 VLOOKUP / SUMPRODUCT / 合计类公式（与 Inp_Hg content、DF_NFMO 联动）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.1 IND_NFMO_NFMS_ZnS.xlsx", "Inp_Matric_HgZn")
