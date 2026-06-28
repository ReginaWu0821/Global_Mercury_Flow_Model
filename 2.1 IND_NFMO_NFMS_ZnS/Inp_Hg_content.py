"""
工作簿: Global_Hg_Flow/2.1 IND_NFMO_NFMS_ZnS.xlsx
工作表: 'Inp_Hg content'
已用区域: A1:I46
合并单元格块数: 2
首行预览:  | Hg content (mg/kg) |  |  |  | Zn content (%) |  |  |  |  |  |  |  |  |  | 

实现说明: 黄色底为输入；在此实现该 sheet 的公式/计算逻辑，可复用 openpyxl / pandas 读源表。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.1 IND_NFMO_NFMS_ZnS.xlsx", "Inp_Hg content")
