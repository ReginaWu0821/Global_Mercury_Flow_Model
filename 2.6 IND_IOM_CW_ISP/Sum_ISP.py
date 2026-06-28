"""
工作簿: Global_Hg_Flow/2.6 IND_IOM_CW_ISP.xlsx
工作表: 'Sum_ISP'
已用区域: A1:S232
合并单元格块数: —
首行预览: Country or Aera | … | ISP 汞流与汇总列 O–S

实现说明: D–S 为公式；引用 Inp_ISP、Sum_IOM_CW、DF_APCD。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.6 IND_IOM_CW_ISP.xlsx", "Sum_ISP")
