"""
工作簿: Global_Hg_Flow/2.7 IND_CEM.xlsx
工作表: 'Sum_CEM'
已用区域: A1:X232
合并单元格块数: —
首行预览: Country or Aera | … | 排放分配与占比 S/V/X

实现说明: D–M、N–R、S/V/X 为公式；引用 Inp_CEM、DF_APCD；占比行引用 Sum_CEM 自身块 N:R。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.7 IND_CEM.xlsx", "Sum_CEM")
