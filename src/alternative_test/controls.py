"""余额—发生额数学勾稽。"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from enum import Enum

from .models import CounterpartyBalance, NormalBalance
from .population import PopulationResult


class ControlStatus(str, Enum):
    PASS = "PASS"
    DIFFERENCE = "DIFFERENCE"
    NOT_RUN = "NOT_RUN"
    NEEDS_INPUT = "NEEDS_INPUT"


@dataclass
class ControlResult:
    """单个供应商的勾稽结果。"""
    counterparty_id: str
    raw_name: str
    reported_opening: Decimal | None
    reported_ending: Decimal
    total_debit: Decimal = Decimal("0")
    total_credit: Decimal = Decimal("0")
    expected_ending: Decimal | None = None
    difference: Decimal | None = None
    status: ControlStatus = ControlStatus.NOT_RUN
    tolerance: Decimal = Decimal("0.00")
    decision_id: str | None = None


def reconcile_counterparty(
    balance: CounterpartyBalance,
    population: PopulationResult,
    tolerance: Decimal = Decimal("0.00"),
) -> ControlResult:
    """对单个供应商执行期初+发生=期末勾稽。

    借方正常余额：expected_ending = opening + debit - credit
    贷方正常余额：expected_ending = opening + credit - debit
    """
    result = ControlResult(
        counterparty_id=balance.counterparty_id,
        raw_name=balance.raw_name,
        reported_opening=balance.opening_balance,
        reported_ending=balance.ending_balance,
        total_debit=population.total_debit,
        total_credit=population.total_credit,
        tolerance=tolerance,
    )

    # 期初缺失 → NOT_RUN
    if balance.opening_balance is None:
        result.status = ControlStatus.NOT_RUN
        return result

    # 计算预期期末
    if balance.normal_balance == NormalBalance.DEBIT:
        expected = balance.opening_balance + population.total_debit - population.total_credit
    else:
        expected = balance.opening_balance + population.total_credit - population.total_debit

    expected = expected  # 保留 Decimal 精度
    diff = balance.ending_balance - expected

    result.expected_ending = expected
    result.difference = diff

    if abs(diff) <= tolerance:
        result.status = ControlStatus.PASS
    else:
        result.status = ControlStatus.DIFFERENCE

    return result


def run_controls(
    balances: list[CounterpartyBalance],
    populations: dict[str, PopulationResult],  # counterparty_id → population
    tolerance: Decimal = Decimal("0.00"),
) -> list[ControlResult]:
    """批量执行勾稽。"""
    results: list[ControlResult] = []
    for b in balances:
        pop = populations.get(b.counterparty_id, PopulationResult())
        results.append(reconcile_counterparty(b, pop, tolerance))
    return results
