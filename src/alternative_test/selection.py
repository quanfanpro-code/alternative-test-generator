"""用户规则驱动的候选选择 — 无规则时保持完整总体。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Literal

from pydantic import BaseModel

from .models import JournalLine, NormalBalance
from .population import classify_movement
from .population import PopulationResult

# ──────────────────────────────────────────────
# 选择规则模型
# ──────────────────────────────────────────────

class SelectionRule(BaseModel):
    """候选选择规则。"""
    model_config = {"extra": "forbid"}

    rule_id: str
    field: Literal[
        "signed_amount",
        "absolute_amount",
        "posting_date",
        "voucher_key",
        "movement_kind",
    ]
    operator: Literal["gte", "lte", "between", "equals", "in"]
    value: Decimal | date | str | list[str]
    user_confirmed: bool = False
    description: str = ""


@dataclass
class SelectionResult:
    """选择结果。"""
    population: list[JournalLine] = field(default_factory=list)
    selected: list[tuple[JournalLine, str]] = field(default_factory=list)  # (line, rule_id)
    excluded: list[tuple[JournalLine, str]] = field(default_factory=list)


# ──────────────────────────────────────────────
# select_candidates
# ──────────────────────────────────────────────

def select_candidates(
    population: PopulationResult,
    rules: list[SelectionRule],
    normal_balance: NormalBalance = NormalBalance.CREDIT,
) -> SelectionResult:
    """根据用户规则从总体中选择候选。

    - 无规则或规则未确认 → 返回全量 population，selected 为空
    - 只有 user_confirmed=True 的规则才执行
    """
    result = SelectionResult(population=list(population.included))

    # 过滤出已确认的规则
    active_rules = [r for r in rules if r.user_confirmed]
    if not active_rules:
        return result  # 无规则 → 保持完整总体，selected 为空

    # 对每条 line 检查是否命中任一规则
    for line in population.included:
        hit = False
        for rule in active_rules:
            if _check_rule(line, rule, normal_balance):
                result.selected.append((line, rule.rule_id))
                hit = True
                break  # 命中一条规则即可
        if not hit:
            result.excluded.append((line, "未命中任何规则"))

    return result


def _check_rule(line: JournalLine, rule: SelectionRule, normal_balance=NormalBalance.CREDIT) -> bool:
    """检查行是否命中单条规则。"""
    field_value = _get_field_value(line, rule.field, normal_balance)
    values = rule.value if isinstance(rule.value, list) else [rule.value]
    if isinstance(field_value, Decimal):
        values = [Decimal(str(v)) for v in values]
    elif isinstance(field_value, date):
        values = [v if isinstance(v, date) else date.fromisoformat(str(v)) for v in values]
    value = values if isinstance(rule.value, list) else values[0]

    if rule.operator == "gte":
        return field_value >= value

    elif rule.operator == "lte":
        return field_value <= value

    elif rule.operator == "between" and isinstance(value, list) and len(value) == 2:
        return value[0] <= field_value <= value[1]

    elif rule.operator == "equals":
        return field_value == value

    elif rule.operator == "in" and isinstance(value, list):
        return field_value in value

    return False


def _get_field_value(line: JournalLine, field: str, normal_balance=NormalBalance.CREDIT):
    """获取规则字段值。"""
    if field == "signed_amount":
        # 使用绝对值较大的方向
        return line.debit if abs(line.debit) >= abs(line.credit) else line.credit
    elif field == "absolute_amount":
        return max(abs(line.debit), abs(line.credit))
    elif field == "posting_date":
        return line.posting_date
    elif field == "voucher_key":
        return str(line.voucher_key)
    elif field == "movement_kind":
        return classify_movement(line, normal_balance).value
    return None
