"""
工作簿: Global_Hg_Flow/1.1 COMCoal_CW_CFPP.xlsx
工作表: 'DF_Wastes'
已用区域: A1:AI94
合并单元格块数: 11
首行预览:  | Waste type | Country or Aera | Utilization Proportion (%) | Utilization Pathways (%) |  |  |  |  |  |  | Distribution Factor |  |  |  |  | 

实现说明: 黄色底为输入；在此实现该 sheet 的公式/计算逻辑，可复用 openpyxl / pandas 读源表。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.1 COMCoal_CW_CFPP.xlsx", "DF_Wastes")
