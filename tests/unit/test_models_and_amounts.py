"""Task 2 — Decimal 金额解析、模型序列化与 ID 稳定性测试。"""

from __future__ import annotations

import json
from datetime import date
from decimal import Decimal

import pytest

from alternative_test.amounts import parse_amount, quantize_money
from alternative_test.models import (
    CounterpartyBalance,
    DecisionStatus,
    JournalLine,
    MatchDecision,
    NormalBalance,
    SourceRef,
)

# ──────────────────────────────────────────────
# parse_amount 参数化测试
# ──────────────────────────────────────────────

def _dummy_source() -> SourceRef:
    return SourceRef(
        file_id="test.xlsx",
        sha256="a" * 64,
        sheet="Sheet1",
        row=1,
        columns={"amount": "H"},
    )


@pytest.mark.parametrize(
    "value,expected",
    [
        (1234.56, Decimal("1234.56")),
        ("1,234.56", Decimal("1234.56")),
        ("-1234.56", Decimal("-1234.56")),
        ("(1,234.56)", Decimal("-1234.56")),
        (1234, Decimal("1234")),
        ("0", Decimal("0")),
        ("0.00", Decimal("0.00")),
        (Decimal("9876.54"), Decimal("9876.54")),
        (None, None),
        ("", None),
        ("   ", None),
    ],
)
def test_parse_amount_valid(value, expected) -> None:
    """正常金额可解析为 Decimal；空白/None 返回 None。"""
    result = parse_amount(value, source=_dummy_source(), field="debit")
    assert result == expected


def test_parse_amount_illegal_raises() -> None:
    """非法字符串金额抛出 AmountParseError，禁止静默 0。"""
    from alternative_test.errors import AmountParseError

    with pytest.raises(AmountParseError):
        parse_amount("not_a_number", source=_dummy_source(), field="debit")


def test_parse_amount_red_text() -> None:
    """红字文本（如 '红字 100' 或 '冲销 200'）应返回原文让上层处理，不能静默转 0。"""
    # 红字/冲销文本无法直接解析，应抛异常
    from alternative_test.errors import AmountParseError

    with pytest.raises(AmountParseError):
        parse_amount("红字 100.00", source=_dummy_source(), field="debit")


# ──────────────────────────────────────────────
# quantize_money
# ──────────────────────────────────────────────

def test_quantize_money_rounds_to_cents() -> None:
    """金额量化到分（两位小数）。"""
    assert quantize_money(Decimal("100.005")) == Decimal("100.01")


# ──────────────────────────────────────────────
# Decimal 精确性
# ──────────────────────────────────────────────

def test_decimal_precision() -> None:
    """0.1 + 0.2 == 0.3（Decimal 下精确成立）。"""
    a = Decimal("0.1")
    b = Decimal("0.2")
    assert a + b == Decimal("0.3")


# ──────────────────────────────────────────────
# 模型构造与序列化
# ──────────────────────────────────────────────

def test_source_ref_required_fields() -> None:
    """SourceRef 必填字段校验。"""
    sr = SourceRef(
        file_id="f1.xlsx",
        sha256="a" * 64,
        sheet="余额表",
        row=5,
        columns={"opening": "C", "ending": "F"},
    )
    assert sr.file_id == "f1.xlsx"
    assert len(sr.sha256) == 64


def test_journal_line_voucher_key() -> None:
    """凭证键 = (entity, ledger, posting_date, voucher_type, voucher_number)。"""
    jl = JournalLine(
        entity_id="E001",
        ledger_id="L001",
        posting_date=date(2025, 6, 30),
        voucher_type="记账凭证",
        voucher_number="24.0001",
        line_number=1,
        account_code="2202",
        account_name="应付账款",
        counterparty_raw="合成采购有限公司",
        summary="采购原材料",
        debit=Decimal("10000.00"),
        credit=Decimal("0.00"),
        source=_dummy_source(),
    )
    vk = jl.voucher_key
    assert vk == ("E001", "L001", date(2025, 6, 30), "记账凭证", "24.0001")


def test_counterparty_balance_normal_balance() -> None:
    """CounterpartyBalance 含正常余额方向。"""
    cb = CounterpartyBalance(
        entity_id="E001",
        counterparty_id="CP001",
        raw_name="合成采购有限公司",
        account_code="FE001",
        account_name="应付账款",
        normal_balance=NormalBalance.CREDIT,
        opening_balance=Decimal("50000.00"),
        ending_balance=Decimal("80000.00"),
        source=_dummy_source(),
    )
    assert cb.normal_balance == NormalBalance.CREDIT
    assert cb.opening_balance == Decimal("50000.00")


def test_json_roundtrip_decimal_negative() -> None:
    """负数 Decimal 可 JSON 往返。"""
    jl = JournalLine(
        entity_id="E001",
        ledger_id="L001",
        posting_date=date(2025, 3, 15),
        voucher_type="记账凭证",
        voucher_number="12.0001",
        line_number=1,
        account_code="2202",
        account_name="应付账款",
        counterparty_raw="合成采购有限公司",
        summary="退货冲销",
        debit=Decimal("-5000.00"),
        credit=Decimal("0.00"),
        source=_dummy_source(),
    )
    dumped = jl.model_dump_json()
    loaded = json.loads(dumped)
    assert loaded["debit"] == "-5000.00"
    assert loaded["credit"] == "0.00"
    # 往返还原
    restored = JournalLine.model_validate_json(dumped)
    assert restored.debit == Decimal("-5000.00")


def test_extra_fields_forbidden() -> None:
    """模型 extra='forbid' — 未知字段拒绝。"""
    with pytest.raises(Exception):
        SourceRef(**{
            "file_id": "f1",
            "sha256": "a" * 64,
            "sheet": "S1",
            "row": 1,
            "columns": {},
            "unknown_field": "should_fail",
        })


def test_decision_status_enum() -> None:
    """DecisionStatus 枚举值正确。"""
    assert DecisionStatus.MISSING.value == "missing"
    assert DecisionStatus.CONFIRMED.value == "user_confirmed"
    assert DecisionStatus.UNAVAILABLE.value == "unavailable"


def test_normal_balance_enum() -> None:
    """NormalBalance 枚举值正确。"""
    assert NormalBalance.DEBIT.value == "debit"
    assert NormalBalance.CREDIT.value == "credit"


def test_match_decision_status() -> None:
    """MatchDecision 状态字段与候选 ID 校验。"""
    md = MatchDecision(
        journal_line_id="JL001",
        counterparty_id="CP001",
        status="exact_unique",
        candidate_ids=["CP001"],
        reason="全字符串精确匹配，且唯一",
        decision_id="D001",
    )
    assert md.status == "exact_unique"
    assert md.counterparty_id == "CP001"
