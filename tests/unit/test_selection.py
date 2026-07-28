"""Task 8 — 候选选择规则测试。"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from alternative_test.models import JournalLine, SourceRef
from alternative_test.population import PopulationResult
from alternative_test.selection import (
    SelectionRule,
    select_candidates,
)


def _source() -> SourceRef:
    return SourceRef(file_id="t.xlsx", sha256="a" * 64, sheet="S", row=1, columns={})


def _line(
    dt: date = date(2025, 6, 30),
    debit: Decimal = Decimal("0"),
    credit: Decimal = Decimal("0"),
) -> JournalLine:
    return JournalLine(
        entity_id="E001", ledger_id="L001", posting_date=dt,
        voucher_type="记账凭证", voucher_number="24.0001", line_number=1,
        account_code="2202", account_name="应付账款",
        counterparty_raw="合成采购有限公司", summary="采购",
        debit=debit, credit=credit, source=_source(),
    )


# ── 无规则 → 完整总体 ──

def test_no_rules_returns_full_population() -> None:
    """无规则时返回完整总体，selected 为空。"""
    pop = PopulationResult(included=[
        _line(debit=Decimal("10000")),
        _line(debit=Decimal("5000")),
    ])
    result = select_candidates(pop, [])
    assert len(result.population) == 2
    assert len(result.selected) == 0


def test_unconfirmed_rule_not_applied() -> None:
    """未确认规则不执行，selected 为空。"""
    pop = PopulationResult(included=[
        _line(debit=Decimal("10000")),
        _line(debit=Decimal("5000")),
    ])
    rule = SelectionRule(
        rule_id="R1",
        field="absolute_amount",
        operator="gte",
        value=Decimal("8000"),
        user_confirmed=False,  # 未确认
    )
    result = select_candidates(pop, [rule])
    assert len(result.selected) == 0


# ── 已确认规则 ──

def test_amount_threshold_rule() -> None:
    """金额门槛规则命中。"""
    pop = PopulationResult(included=[
        _line(debit=Decimal("10000")),
        _line(debit=Decimal("2000")),
    ])
    rule = SelectionRule(
        rule_id="R1",
        field="absolute_amount",
        operator="gte",
        value=Decimal("5000"),
        user_confirmed=True,
    )
    result = select_candidates(pop, [rule])
    assert len(result.selected) == 1
    line, rid = result.selected[0]
    assert line.debit == Decimal("10000")
    assert rid == "R1"


def test_multiple_rules() -> None:
    """多规则：命中任一即选中。"""
    pop = PopulationResult(included=[
        _line(debit=Decimal("10000"), dt=date(2025, 1, 1)),
        _line(debit=Decimal("2000"), dt=date(2025, 12, 31)),
    ])
    rules = [
        SelectionRule(
            rule_id="R1", field="absolute_amount", operator="gte",
            value=Decimal("5000"), user_confirmed=True,
        ),
        SelectionRule(
            rule_id="R2", field="posting_date", operator="equals",
            value=date(2025, 12, 31), user_confirmed=True,
        ),
    ]
    result = select_candidates(pop, rules)
    # 第二条规则也命中小额行
    assert len(result.selected) == 2


def test_amount_not_changed_by_rule() -> None:
    """规则选择不改变金额。"""
    pop = PopulationResult(included=[
        _line(debit=Decimal("-5000")),  # 红字
    ])
    rule = SelectionRule(
        rule_id="R1",
        field="absolute_amount",
        operator="gte",
        value=Decimal("1000"),
        user_confirmed=True,
    )
    result = select_candidates(pop, [rule])
    assert len(result.selected) == 1
    line, _ = result.selected[0]
    # 金额保持原负值
    assert line.debit == Decimal("-5000")
