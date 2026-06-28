"""
工作簿: Global_Hg_Flow/2.1 IND_NFMO_NFMS_ZnS.xlsx
工作表: 'Sum_NFMO_Zn'
已用区域: A1:I232
合并单元格块数: 5
首行预览: Country or Aera | Abbrev | Abbrev | Hg input | HgFlow_Pre |  |  |  |  |  |  |  |  |  |  | 

实现说明: A–C、表头为输入；D–I 为按 Inp_ZnS 汞输入与 DF_NFMO 分配的行公式及合计。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.1 IND_NFMO_NFMS_ZnS.xlsx", "Sum_NFMO_Zn")
