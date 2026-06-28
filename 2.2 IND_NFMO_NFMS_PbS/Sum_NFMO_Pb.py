"""
工作簿: Global_Hg_Flow/2.2 IND_NFMO_NFMS_PbS.xlsx
工作表: 'Sum_NFMO_Pb'
已用区域: A1:I232
合并单元格块数: 5
首行预览: Country or Aera | Abbrev | Abbrev | Hg input | HgFlow_Pre |  |  |  |  |  |  |  |  |  |  | 

实现说明: A–C 为输入；D–I 按 Inp_PbS 汞输入与 DF_NFMO 分配。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.2 IND_NFMO_NFMS_PbS.xlsx", "Sum_NFMO_Pb")
