"""
工作簿: Global_Hg_Flow/2.4 IND_NFMO_NFMS_LSGP.xlsx
工作表: 'Sum_LSGP'
已用区域: A1:P436
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev for calculation | Abbrev | Region | Sub-region | Emission and distribution (t) |  |  |  |  |  |  |  |  | 

实现说明: A–F、C 列为输入；G–M 为 K/L 汞输入与 DF_LSGP（按 B 列匹配及第 11 行固定项）的组合公式；公式行为第 4–232 行。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.4 IND_NFMO_NFMS_LSGP.xlsx", "Sum_LSGP")
