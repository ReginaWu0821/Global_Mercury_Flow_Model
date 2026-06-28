"""
工作簿: Global_Hg_Flow/2.5 IND_NFMO_AlS.xlsx
工作表: 'Sum_AlS'
已用区域: A1:N436
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Emission and distribution (t) |  |  |  |  |  |  |  |  | 

实现说明: G–K 为 Inp_AlS 汞输入 F 与分配因子列 M–Q 的乘积；第 3 行为列合计。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.5 IND_NFMO_AlS.xlsx", "Sum_AlS")
