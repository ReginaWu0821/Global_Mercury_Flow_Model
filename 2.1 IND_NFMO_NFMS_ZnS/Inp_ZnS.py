"""
工作簿: Global_Hg_Flow/2.1 IND_NFMO_NFMS_ZnS.xlsx
工作表: 'Inp_ZnS'
已用区域: A1:N233
合并单元格块数: 12
首行预览: Country or Aera | Abbrev | AL_Loc (kt) | AL_Imp (kt) | AL_Production (kt) | AL_Loc_Pro (%) | AL_Imp_Pro (%) | HgCon_Loc (mg/kg) | HgCon_Imp (mg/kg) | ZnCon_Loc (%) | ZnCon_Imp (%) | HgInp (t) |  |  |  | 

实现说明: A–B、E 及活动/含量相关列为输入或中间结构；C、D、F–N 为与矩阵及 DF_NFMO 相关的公式。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.1 IND_NFMO_NFMS_ZnS.xlsx", "Inp_ZnS")
