"""Hg_flow：按模板行展开 Excel 公式（与 Global_Hg_Flow 工作簿生成脚本配套）。"""

from __future__ import annotations

import re


def expand_formula(formula: str, src: int, dst: int) -> str:
    """将模板行 ``src`` 上的公式展开到目标行 ``dst``。

    所有"跟在列字母后的整数行号"都按相对偏移平移：把行号 ``n`` 替换为
    ``dst + (n - src)``。形式上识别以下四类引用：

    - ``COLUMN(Letters{n})``
    - ``$Col{n}``（列绝对、行相对；常见于 ``$A5`` 这种）
    - ``!Col{n}``（外部/跨表引用，如 ``Sheet!B4``）
    - 裸单元格 ``Col{n}``（不带 ``$``、不带 ``!``、不在 ``[N]`` 工作簿索引里）

    被 ``$`` 锁住的整数（如 ``$C$4``、``$B$23:$Q$30``）属于绝对锚点，不会被替换；
    工作簿索引 ``[3]``、``[11]`` 因为前导是 ``[``，也不会被命中。

    ``COLUMN(...)`` 在替换前先换成占位符，避免 ``COLUMN(J5)`` 中的 ``J5`` 被
    "裸引用"规则误替换。
    """

    placeholders: list[str] = []
    delta = dst - src

    def col_repl(m: re.Match[str]) -> str:
        letters = m.group(1)
        new_inner = int(m.group(2)) + delta
        token = f"__HG_COLPH_{len(placeholders)}__"
        placeholders.append(f"COLUMN({letters}{new_inner})")
        return token

    s = re.sub(r"COLUMN\(([A-Z]+)(\d+)\)", col_repl, formula)
    # ``$Col{n}`` -> 行号 +delta（注意 ``$Col$n`` 不会被命中，因为 ``$`` 后是字母）
    s = re.sub(
        r"\$([A-Z]+)(\d+)\b",
        lambda m: f"${m.group(1)}{int(m.group(2)) + delta}",
        s,
    )
    # ``!Col{n}`` -> 行号 +delta
    s = re.sub(
        r"!([A-Z]+)(\d+)\b",
        lambda m: f"!{m.group(1)}{int(m.group(2)) + delta}",
        s,
    )
    # 裸 ``Col{n}`` -> 行号 +delta（前导不是 ``$`` 也不是 ``!``）
    s = re.sub(
        r"(?<![\$!])\b([A-Z]+)(\d+)\b",
        lambda m: f"{m.group(1)}{int(m.group(2)) + delta}",
        s,
    )
    for i, ph in enumerate(placeholders):
        s = s.replace(f"__HG_COLPH_{i}__", ph)
    return s


def refill_from_template_row(
    ws,
    template_row: int,
    first_row: int,
    last_row: int,
    *,
    preserve_row3: bool = True,
) -> None:
    """将 ``template_row`` 上的公式列复制展开到 ``first_row:last_row``（含端点）。

    模板行在写入前读取；若 ``preserve_row3`` 为真，则先快照第 3 行并在写入后还原，
    以免覆盖合计行或比例列等特殊公式。
    """

    max_col = ws.max_column
    row3_snapshot = None
    if preserve_row3:
        row3_snapshot = [ws.cell(3, c).value for c in range(1, max_col + 1)]
    templates = [ws.cell(template_row, c).value for c in range(1, max_col + 1)]
    for dst in range(first_row, last_row + 1):
        for c in range(1, max_col + 1):
            tpl = templates[c - 1]
            if isinstance(tpl, str) and tpl.startswith("="):
                ws.cell(dst, c).value = expand_formula(tpl, template_row, dst)
    if preserve_row3 and row3_snapshot is not None:
        for c in range(1, max_col + 1):
            ws.cell(3, c).value = row3_snapshot[c - 1]
