"""Task 7 — 余额勾稽测试。"""

from __future__ import annotations

from decimal import Decimal

from alternative_test.controls import (
    ControlStatus,
    reconcile_counterparty,
    run_controls,
)
from alternative_test.models import CounterpartyBalance, NormalBalance, SourceRef
from alternative_test.population import PopulationResult


def _source() -> SourceRef:
    return SourceRef(file_id="t.xlsx", sha256="a" * 64, sheet="S", row=1, columns={})


def _balance(opening: Decimal | None, ending: Decimal) -> CounterpartyBalance:
    return CounterpartyBalance(
        entity_id="E001",
        counterparty_id="CP001",
        raw_name="合成采购有限公司",
        account_code="FE001",
        account_name="应付账款",
        normal_balance=NormalBalance.CREDIT,
        opening_balance=opening,
        ending_balance=ending,
        source=_source(),
    )


# ── reconcile_counterparty ──

def test_reconcile_pass() -> None:
    """期初+贷方-借方=期末 → PASS。"""
    b = _balance(Decimal("10000"), Decimal("25000"))
    pop = PopulationResult(
        total_debit=Decimal("5000"),
        total_credit=Decimal("20000"),
    )
    result = reconcile_counterparty(b, pop)
    # expected = 10000 + 20000 - 5000 = 25000
    assert result.status == ControlStatus.PASS
    assert result.expected_ending == Decimal("25000")
    assert result.difference == Decimal("0")


def test_reconcile_difference() -> None:
    """存在差异 → DIFFERENCE。"""
    b = _balance(Decimal("10000"), Decimal("30000"))
    pop = PopulationResult(
        total_debit=Decimal("5000"),
        total_credit=Decimal("20000"),
    )
    result = reconcile_counterparty(b, pop)
    # expected = 25000, actual = 30000, diff = 5000
    assert result.status == ControlStatus.DIFFERENCE
    assert result.difference == Decimal("5000")


def test_reconcile_opening_missing() -> None:
    """期初缺失 → NOT_RUN。"""
    b = _balance(None, Decimal("25000"))
    pop = PopulationResult(
        total_debit=Decimal("5000"),
        total_credit=Decimal("20000"),
    )
    result = reconcile_counterparty(b, pop)
    assert result.status == ControlStatus.NOT_RUN
    assert result.expected_ending is None


def test_reconcile_debit_normal() -> None:
    """借方正常余额：expected = opening + debit - credit。"""
    b = CounterpartyBalance(
        entity_id="E001",
        counterparty_id="CP001",
        raw_name="预付供应商",
        account_code="ZG001",
        account_name="预付账款",
        normal_balance=NormalBalance.DEBIT,
        opening_balance=Decimal("50000"),
        ending_balance=Decimal("90000"),
        source=_source(),
    )
    pop = PopulationResult(
        total_debit=Decimal("60000"),
        total_credit=Decimal("20000"),
    )
    result = reconcile_counterparty(b, pop)
    # expected = 50000 + 60000 - 20000 = 90000
    assert result.status == ControlStatus.PASS
    assert result.expected_ending == Decimal("90000")


def test_reconcile_red_ink() -> None:
    """红字/负数参与勾稽。"""
    b = _balance(Decimal("10000"), Decimal("5000"))
    pop = PopulationResult(
        total_debit=Decimal("-1000"),  # 红字借方
        total_credit=Decimal("-4000"),  # 红字贷方
    )
    result = reconcile_counterparty(b, pop)
    # expected = 10000 + (-4000) - (-1000) = 10000 - 4000 + 1000 = 7000
    assert result.expected_ending == Decimal("7000")
    assert result.difference == Decimal("-2000")


def test_reconcile_tolerance() -> None:
    """容差范围内视为 PASS。"""
    b = _balance(Decimal("10000"), Decimal("25000.01"))
    pop = PopulationResult(
        total_debit=Decimal("5000"),
        total_credit=Decimal("20000"),
    )
    result = reconcile_counterparty(b, pop, tolerance=Decimal("0.02"))
    assert result.status == ControlStatus.PASS


# ── run_controls ──

def test_run_controls_multiple() -> None:
    """批量勾稽多个供应商。"""
    b1 = _balance(Decimal("10000"), Decimal("25000"))
    b2 = _balance(None, Decimal("50000"))  # 期初缺失
    pop = PopulationResult(
        total_debit=Decimal("5000"),
        total_credit=Decimal("20000"),
    )
    results = run_controls(
        [b1, b2],
        {"CP001": pop, "CP002": pop},
    )
    assert len(results) == 2
    assert results[0].status == ControlStatus.PASS
    assert results[1].status == ControlStatus.NOT_RUN
