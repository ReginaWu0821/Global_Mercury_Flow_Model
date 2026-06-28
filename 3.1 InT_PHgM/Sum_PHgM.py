"""
工作簿: Global_Hg_Flow/3.1 InT_PHgM.xlsx
工作表: 'Sum_PHgM'
已用区域: A1:L232
合并单元格块数: 3
首行预览:  | Hg Input (t) | Distribution Factor |  |  |  |  | Emissin and Releases |  |  |  |  |  |  |  | 

实现说明: 绝大多数国家行仅有国别输入；含汞产量与分配因子（C–G、L）为原始/外部输入。
仅在若干指定行写入物料衡算公式（中国单独形式；印尼等为 C=1−SUM(D:G) 形式）。
"""

import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO / "Code"))
from _sheet_formula_sync import sync_complex_formulas

sync_complex_formulas(_REPO, "3.1 InT_PHgM.xlsx", "Sum_PHgM")
