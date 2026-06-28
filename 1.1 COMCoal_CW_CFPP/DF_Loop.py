"""
工作簿: Global_Hg_Flow/1.1 COMCoal_CW_CFPP.xlsx
工作表: 'DF_Loop'
已用区域: A1:BB232
合并单元格块数: 12
首行预览:  | Country or Aera | Abbrev | Wst df_fly ash | Wst df_gypsum | EU27 | Reg_code | Flw_Wst_In |  |  | Fly ash |  |  |  |  |  | 

实现说明: 黄色底为输入；在此实现该 sheet 的公式/计算逻辑，可复用 openpyxl / pandas 读源表。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "1.1 COMCoal_CW_CFPP.xlsx", "DF_Loop")
