"""Task 6 — 期间过滤、凭证键、科目方向测试。"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from alternative_test.matching import MatchPlan
from alternative_test.models import JournalLine, NormalBalance, SourceRef
from alternative_test.population import (
    ConfirmedScope,
    MovementKind,
    build_population,
    classify_movement,
)


def _source() -> SourceRef:
    return SourceRef(file_id="t.xlsx", sha256="a" * 64, sheet="S", row=1, columns={})


def _line(
    entity: str = "E001",
    ledger: str = "L001",
    dt: date = date(2025, 6, 30),
    vtype: str = "记账凭证",
    vno: str = "24.0001",
    debit: Decimal = Decimal("0"),
    credit: Decimal = Decimal("0"),
) -> JournalLine:
    return JournalLine(
        entity_id=entity, ledger_id=ledger, posting_date=dt,
        voucher_type=vtype, voucher_number=vno, line_number=1,
        account_code="2202", account_name="应付账款",
        counterparty_raw="合成采购有限公司", summary="采购",
        debit=debit, credit=credit, source=_source(),
    )


# ── classify_movement ──

def test_credit_normal_balance_credit_is_increase() -> None:
    """贷方正常余额下 credit 为增加。"""
    line = _line(debit=Decimal("0"), credit=Decimal("5000"))
    assert classify_movement(line, NormalBalance.CREDIT) == MovementKind.INCREASE


def test_credit_normal_balance_debit_is_decrease() -> None:
    """贷方正常余额下 debit 为减少。"""
    line = _line(debit=Decimal("5000"), credit=Decimal("0"))
    assert classify_movement(line, NormalBalance.CREDIT) == MovementKind.DECREASE


def test_debit_normal_balance_debit_is_increase() -> None:
    """借方正常余额下 debit 为增加。"""
    line = _line(debit=Decimal("5000"), credit=Decimal("0"))
    assert classify_movement(line, NormalBalance.DEBIT) == MovementKind.INCREASE


def test_debit_normal_balance_credit_is_decrease() -> None:
    """借方正常余额下 credit 为减少。"""
    line = _line(debit=Decimal("0"), credit=Decimal("5000"))
    assert classify_movement(line, NormalBalance.DEBIT) == MovementKind.DECREASE


# ── build_population ──

def test_build_population_period_filter() -> None:
    """期间外记录进入 excluded。"""
    scope = ConfirmedScope(start_date=date(2025, 1, 1), end_date=date(2025, 12, 31))
    lines = [
        _line(dt=date(2025, 6, 1), debit=Decimal("1000")),
        _line(dt=date(2024, 12, 1), debit=Decimal("2000")),  # 期间外
        _line(dt=date(2026, 1, 1), debit=Decimal("3000")),  # 期间外
    ]
    result = build_population(lines, scope, MatchPlan())
    assert len(result.included) == 1
    assert len(result.excluded) == 2


def test_build_population_preserves_negative() -> None:
    """红字/负数保留原符号。"""
    scope = ConfirmedScope(start_date=date(2025, 1, 1), end_date=date(2025, 12, 31))
    lines = [
        _line(dt=date(2025, 6, 1), debit=Decimal("-1000"), credit=Decimal("0")),
    ]
    result = build_population(lines, scope, MatchPlan())
    assert len(result.included) == 1
    assert result.included[0].debit == Decimal("-1000")
    assert result.total_debit == Decimal("-1000")


def test_build_population_entity_filter() -> None:
    """实体范围过滤。"""
    scope = ConfirmedScope(
        start_date=date(2025, 1, 1),
        end_date=date(2025, 12, 31),
        entity_ids=["E001"],
    )
    lines = [
        _line(entity="E001", dt=date(2025, 6, 1)),
        _line(entity="E002", dt=date(2025, 6, 1)),
    ]
    result = build_population(lines, scope, MatchPlan())
    assert len(result.included) == 1


def test_voucher_key_prevents_merge() -> None:
    """跨月同号不合并（不同凭证键）。"""
    line1 = _line(dt=date(2025, 6, 1), vno="001")
    line2 = _line(dt=date(2025, 7, 1), vno="001")
    assert line1.voucher_key != line2.voucher_key
