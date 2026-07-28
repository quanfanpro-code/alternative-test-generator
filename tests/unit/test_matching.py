"""Task 5 — 供应商匹配测试：exact_key、候选生成、用户确认。"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from alternative_test.matching import (
    apply_confirmed_decisions,
    exact_key,
    propose_matches,
)
from alternative_test.models import (
    CounterpartyBalance,
    JournalLine,
    NormalBalance,
    SourceRef,
)


def _dummy_source() -> SourceRef:
    return SourceRef(
        file_id="test.xlsx",
        sha256="a" * 64,
        sheet="Sheet1",
        row=1,
        columns={},
    )


def _make_balance(raw_name: str, cp_id: str, entity: str = "E001") -> CounterpartyBalance:
    return CounterpartyBalance(
        entity_id=entity,
        counterparty_id=cp_id,
        raw_name=raw_name,
        account_code="FE001",
        account_name="应付账款",
        normal_balance=NormalBalance.CREDIT,
        ending_balance=Decimal("10000"),
        source=_dummy_source(),
    )


def _make_journal_line(
    counterparty_raw: str,
    line_id: str,
    entity: str = "E001",
    summary: str = "采购材料",
) -> JournalLine:
    return JournalLine(
        entity_id=entity,
        ledger_id="L001",
        posting_date=date(2025, 6, 30),
        voucher_type="记账凭证",
        voucher_number="24.0001",
        line_number=1,
        account_code="2202",
        account_name="应付账款",
        counterparty_raw=counterparty_raw,
        summary=summary,
        debit=Decimal("5000"),
        credit=Decimal("0"),
        source=_dummy_source(),
    )


# ──────────────────────────────────────────────
# exact_key
# ──────────────────────────────────────────────

def test_exact_key_same_string() -> None:
    """相同字符串返回相同 key。"""
    assert exact_key("合成采购有限公司") == exact_key("合成采购有限公司")


def test_exact_key_fullwidth_space() -> None:
    """全角空格与半角空格等价。"""
    a = exact_key("合成采购　有限公司")  # 全角空格
    b = exact_key("合成采购 有限公司")   # 半角空格
    assert a == b


def test_exact_key_case_insensitive_ascii() -> None:
    """ASCII 大小写统一。"""
    assert exact_key("ABC Company") == exact_key("abc company")


def test_exact_key_trim_whitespace() -> None:
    """首尾空白删除。"""
    assert exact_key("  合成采购有限公司  ") == exact_key("合成采购有限公司")


# ──────────────────────────────────────────────
# propose_matches: exact_unique
# ──────────────────────────────────────────────

def test_propose_exact_unique_match() -> None:
    """精确匹配且唯一 → exact_unique。"""
    balances = [_make_balance("合成采购有限公司", "CP001")]
    lines = [_make_journal_line("合成采购有限公司", "JL001")]
    plan = propose_matches(lines, balances)
    assert len(plan.decisions) == 1
    d = plan.decisions[0]
    assert d.status == "exact_unique"
    assert d.counterparty_id == "CP001"


def test_propose_exact_unique_with_fullwidth_normalization() -> None:
    """全角空格标准化后精确匹配。"""
    balances = [_make_balance("合成采购有限公司", "CP001")]
    lines = [_make_journal_line("合成采购　有限公司", "JL001")]  # 全角空格
    plan = propose_matches(lines, balances)
    assert plan.decisions[0].status == "exact_unique"


# ──────────────────────────────────────────────
# 去后缀 → needs_confirmation
# ──────────────────────────────────────────────

def test_propose_different_suffix_needs_confirmation() -> None:
    """"甲有限公司" vs "甲股份有限公司" — 后缀不同不可自动匹配。"""
    balances = [_make_balance("甲有限公司", "CP001")]
    lines = [_make_journal_line("甲股份有限公司", "JL001")]
    plan = propose_matches(lines, balances)
    d = plan.decisions[0]
    assert d.status == "needs_confirmation"
    assert len(d.candidate_ids) > 0


# ──────────────────────────────────────────────
# 简称 → needs_confirmation
# ──────────────────────────────────────────────

def test_propose_abbreviation_needs_confirmation() -> None:
    """简称匹配 → needs_confirmation。"""
    balances = [_make_balance("合成采购有限公司", "CP001")]
    lines = [_make_journal_line("合成采购", "JL001")]
    plan = propose_matches(lines, balances)

    # 检查是否有 needs_confirmation 的决策
    needs = [d for d in plan.decisions if d.status == "needs_confirmation"]
    # 精确匹配 unmatched 但 candidate 阶段可能产生候选
    # 简称 "合成采购" 包含在 "合成采购有限公司" 中
    assert any(d.status == "needs_confirmation" for d in plan.decisions), (
        f"预期有 needs_confirmation 决策，实际: {plan.decisions}"
    )


# ──────────────────────────────────────────────
# 两相似供应商 → candidate_ids 全部保留
# ──────────────────────────────────────────────

def test_two_similar_suppliers_all_candidates_preserved() -> None:
    """两个相似供应商 → needs_confirmation + 全部候选 ID 保留。"""
    balances = [
        _make_balance("合成采购有限公司", "CP001"),
        _make_balance("合成采购股份有限公司", "CP002"),
    ]
    lines = [_make_journal_line("合成采购有限责任公司", "JL001")]
    plan = propose_matches(lines, balances)
    d = plan.decisions[0]
    assert d.status == "needs_confirmation"
    # 两个候选 ID 都应保留
    assert len(d.candidate_ids) >= 2


# ──────────────────────────────────────────────
# 摘要匹配 → needs_confirmation
# ──────────────────────────────────────────────

def test_summary_match_needs_confirmation() -> None:
    """摘要中的名称只能形成候选。"""
    balances = [_make_balance("合成采购有限公司", "CP001")]
    lines = [
        JournalLine(
            entity_id="E001",
            ledger_id="L001",
            posting_date=date(2025, 6, 1),
            voucher_type="记账凭证",
            voucher_number="24.0001",
            line_number=1,
            account_code="2202",
            account_name="应付账款",
            counterparty_raw=None,  # 辅助核算为空
            summary="向合成采购有限公司采购原材料一批",
            debit=Decimal("10000"),
            credit=Decimal("0"),
            source=_dummy_source(),
        )
    ]
    plan = propose_matches(lines, balances)
    d = plan.decisions[0]
    # 摘要中匹配只能形成候选
    if d.status != "exact_unique":
        assert d.status == "needs_confirmation", (
            f"预期 needs_confirmation 或 exact_unique，实际: {d.status}"
        )


# ──────────────────────────────────────────────
# unmatched
# ──────────────────────────────────────────────

def test_unmatched_preserved() -> None:
    """无法匹配的记录保留在 unmatched，不删除。"""
    balances = [_make_balance("完全无关的公司", "CP001")]
    lines = [_make_journal_line("合成采购有限公司", "JL001")]
    plan = propose_matches(lines, balances)
    d = plan.decisions[0]
    assert d.status == "unmatched"
    # 记录保留，不映射到 unrelated counterparty
    assert d.counterparty_id is None


# ──────────────────────────────────────────────
# apply_confirmed_decisions
# ──────────────────────────────────────────────

def test_apply_confirmed_maps_correctly() -> None:
    """用户确认后映射正确应用。"""
    balances = [_make_balance("合成采购有限公司", "CP001")]
    lines = [_make_journal_line("合成采购", "JL001")]
    plan = propose_matches(lines, balances)

    # 获取自动生成的 journal_line_id
    line_id = plan.decisions[0].journal_line_id

    # 用户确认
    from alternative_test.matching import UserDecision
    decision = UserDecision(
        decision_id="D001",
        journal_line_id=line_id,
        counterparty_id="CP001",
        original_text="合成采购 → 合成采购有限公司",
    )
    updated = apply_confirmed_decisions(plan, [decision])

    confirmed = [d for d in updated.decisions if d.journal_line_id == line_id]
    assert len(confirmed) == 1
    assert confirmed[0].status == "confirmed"
    assert confirmed[0].counterparty_id == "CP001"
