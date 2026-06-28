"""
工作簿: Global_Hg_Flow/2.4 IND_NFMO_NFMS_LSGP.xlsx
工作表: 'Inp_LSGP'
已用区域: A1:M233
合并单元格块数: 7
首行预览: Country or Aera | Abbrev | Activity Level (kt) |  |  | Recover Rate (%) |  | Au content (%) |  | Hg content (g Hg/t Au ore) | HgInp (t) |  |  |  |  | 

实现说明: A–D、F–J 等为原始输入；E 为本地+进口活动量合计；K、L 为分路径汞输入；M 为 K+L。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.4 IND_NFMO_NFMS_LSGP.xlsx", "Inp_LSGP")
