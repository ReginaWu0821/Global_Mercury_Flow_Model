"""
工作簿: Global_Hg_Flow/2.1 IND_NFMO_NFMS_ZnS.xlsx
工作表: 'DF_APCD_ZnS'
已用区域: A1:U60
合并单元格块数: 5
首行预览: Abbrev | APCD | Distribution Factor |  |  |  |  |  |  |  |  | Application Proportion (%) | Distribution Factor for Calculation |  |  | 

实现说明: A–L 为原始输入与说明；M–U 为分布因子 × 应用比例（L 列）/100。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "2.1 IND_NFMO_NFMS_ZnS.xlsx", "DF_APCD_ZnS")
