"""
工作簿: Global_Hg_Flow/2.6 IND_IOM_CW_ISP.xlsx
工作表: 'DF_Loop'
已用区域: A1:Q232
合并单元格块数: 3
首行预览: Country or Aera | TOT (t) | … | 分配比例 K–Q

实现说明: B–J 流量与 VLOOKUP（外部簿 [2]）；K–Q 为 IFERROR 比例，回退行为第 3 行比例。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.6 IND_IOM_CW_ISP.xlsx", "DF_Loop")
