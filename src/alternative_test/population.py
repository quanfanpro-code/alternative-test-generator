"""期间过滤、凭证键、科目方向分类 — 构建候选总体。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from enum import Enum

from .matching import MatchPlan
from .models import JournalLine, NormalBalance


class MovementKind(str, Enum):
    """发生额方向（相对于正常余额）。"""
    INCREASE = "increase"   # 增加正常余额方向的发生额
    DECREASE = "decrease"   # 减少正常余额方向的发生额


@dataclass
class ConfirmedScope:
    """用户确认的范围。"""
    start_date: date
    end_date: date
    entity_ids: list[str] = field(default_factory=list)
    ledger_ids: list[str] = field(default_factory=list)


@dataclass
class PopulationResult:
    """候选总体结果。"""
    included: list[JournalLine] = field(default_factory=list)
    excluded: list[tuple[JournalLine, str]] = field(default_factory=list)
    total_debit: Decimal = Decimal("0")
    total_credit: Decimal = Decimal("0")


def classify_movement(
    line: JournalLine,
    normal_balance: NormalBalance,
) -> MovementKind:
    """根据正常余额方向分类发生额。

    - 贷方正常余额（如应付）：credit 为增加、debit 为减少
    - 借方正常余额（如预付）：debit 为增加、credit 为减少
    """
    if normal_balance == NormalBalance.CREDIT:
        if line.credit != Decimal("0"):
            return MovementKind.INCREASE
        return MovementKind.DECREASE
    else:
        if line.debit != Decimal("0"):
            return MovementKind.INCREASE
        return MovementKind.DECREASE


def build_population(
    lines: list[JournalLine],
    scope: ConfirmedScope,
    matches: MatchPlan,
    normal_balance: NormalBalance = NormalBalance.CREDIT,
) -> PopulationResult:
    """构建候选总体：期间过滤 + 分类。"""
    result = PopulationResult()

    # 构建已匹配行 ID 集合
    matched_ids = {d.journal_line_id for d in matches.decisions
                   if d.status in ("exact_unique", "needs_confirmation", "confirmed")}

    for line in lines:
        # 期间过滤（闭区间）
        if line.posting_date < scope.start_date or line.posting_date > scope.end_date:
            result.excluded.append((line, "期间外"))
            continue

        # 实体过滤
        if scope.entity_ids and line.entity_id not in scope.entity_ids:
            result.excluded.append((line, "实体范围外"))
            continue

        # 账套过滤
        if scope.ledger_ids and line.ledger_id not in scope.ledger_ids:
            result.excluded.append((line, "账套范围外"))
            continue

        # 保留原符号
        result.included.append(line)
        result.total_debit += line.debit
        result.total_credit += line.credit

    return result
