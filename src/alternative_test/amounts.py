"""金额解析与量化 — 全程 Decimal，禁止 abs()。"""

from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from .errors import AmountParseError
from .models import SourceRef

# 金额量化精度：分（两位小数）
_MONEY_SCALE = Decimal("0.01")


def quantize_money(value: Decimal) -> Decimal:
    """将金额量化到分（ROUND_HALF_UP）。"""
    return value.quantize(_MONEY_SCALE, rounding=ROUND_HALF_UP)


def parse_amount(
    value: object,
    *,
    source: SourceRef,
    field: str,
) -> Decimal | None:
    """将任意输入解析为 Decimal 金额。

    - None/空白 → None
    - Decimal → 直接返回
    - int/float → 转为 Decimal
    - str → 去除千分位、处理括号负数、前后负号
    - 非法字符串 → 抛出 AmountParseError（禁止静默转 0）
    """
    if value is None:
        return None

    if isinstance(value, Decimal):
        return value

    if isinstance(value, (int, float)):
        # 浮点转 Decimal 时先转字符串避免二进制误差
        return Decimal(str(value))

    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None

        # 括号负数: "(1,234.56)" → "-1234.56"
        bracket_match = re.match(r"^\((.+)\)$", stripped)
        if bracket_match:
            stripped = bracket_match.group(1)
            negative = True
        else:
            negative = False

        # 去除千分位逗号
        cleaned = stripped.replace(",", "")

        # 处理前后负号
        if cleaned.startswith("-"):
            negative = True
            cleaned = cleaned[1:]
        elif cleaned.endswith("-"):
            negative = True
            cleaned = cleaned[:-1]

        # 再次 strip 去空白
        cleaned = cleaned.strip()

        try:
            result = Decimal(cleaned)
        except (InvalidOperation, ValueError):
            raise AmountParseError(
                f"无法解析金额 '{value}'（field={field}, "
                f"file={source.file_id}, sheet={source.sheet}, row={source.row}）",
                source_info=(
                    f"file={source.file_id}, sheet={source.sheet}, "
                    f"row={source.row}, field={field}"
                ),
            )

        if negative:
            result = -result

        return result

    # 未知类型
    raise AmountParseError(
        f"不支持的金额类型 {type(value).__name__}: {value!r}",
        source_info=f"file={source.file_id}, field={field}",
    )
