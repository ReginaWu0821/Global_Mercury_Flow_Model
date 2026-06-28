"""
工作簿: Global_Hg_Flow/2.5 IND_NFMO_AlS.xlsx
工作表: 'Inp_AlS'
已用区域: A1:Q233
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Abbrev | Activity Level (kt) | Hg Content | HgInp (t) | Production of Alumina from Bauxite |  |  |  | Air pollution Control |  | Distribution Factors |  |  | 

实现说明: E/F/M–Q 为公式；F3 为列合计；外部簿 [1]含汞量 用于含量查找。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.5 IND_NFMO_AlS.xlsx", "Inp_AlS")
