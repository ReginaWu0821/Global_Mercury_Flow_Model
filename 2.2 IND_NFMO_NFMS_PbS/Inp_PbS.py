"""
工作簿: Global_Hg_Flow/2.2 IND_NFMO_NFMS_PbS.xlsx
工作表: 'Inp_PbS'
已用区域: A1:N233
合并单元格块数: 12
首行预览: Country or Aera | Abbrev | AL_Loc (kt) | AL_Imp (kt) | AL_Production (kt) | AL_Loc_Pro (%) | AL_Imp_Pro (%) | HgCon_Loc (mg/kg) | HgCon_Imp (mg/kg) | ZnCon_Loc (%) | ZnCon_Imp (%) | HgInp (t) |  |  |  | 

实现说明: 与 Zn 类似；D 列不含 DF_NFMO×0.95（该因子在矩阵 GZ）；H 列为「零则退回默认」型 VLOOKUP。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.2 IND_NFMO_NFMS_PbS.xlsx", "Inp_PbS")
