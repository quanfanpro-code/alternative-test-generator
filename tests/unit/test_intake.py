"""Task 4 — 角色评分、列映射和逐项追问测试。"""

from __future__ import annotations

from alternative_test.adapters import ColumnMapping
from alternative_test.fileio import TableCandidate
from alternative_test.intake import (
    Answers,
    build_question_queue,
    classify_tables,
    confirm_mapping,
)
from alternative_test.models import DecisionStatus

# ──────────────────────────────────────────────
# classify_tables
# ──────────────────────────────────────────────

def test_classify_balance_table_high_confidence() -> None:
    """余额表含关键词时高置信度识别。"""
    candidates = [
        TableCandidate(
            name="供应商余额表",
            headers=["公司名称", "供应商名称", "科目", "期初余额", "期末余额"],
            row_count=10,
        ),
    ]
    scores = classify_tables(candidates)
    assert len(scores) == 1
    assert scores[0].role == "balance"
    assert scores[0].confidence >= 0.80


def test_classify_journal_table_high_confidence() -> None:
    """序时账含关键词时高置信度识别。"""
    candidates = [
        TableCandidate(
            name="序时账",
            headers=["日期", "凭证类型", "凭证号", "摘要", "借方", "贷方"],
            row_count=100,
        ),
    ]
    scores = classify_tables(candidates)
    assert len(scores) == 1
    assert scores[0].role == "journal"
    assert scores[0].confidence >= 0.80


def test_classify_unknown_table() -> None:
    """无法识别的表标记为 unknown。"""
    candidates = [
        TableCandidate(
            name="随意命名",
            headers=["列A", "列B", "列C"],
            row_count=5,
        ),
    ]
    scores = classify_tables(candidates)
    assert scores[0].role == "unknown"
    assert scores[0].confidence < 0.80


def test_classify_two_equal_score_sheets() -> None:
    """两个同分 Sheet 同时返回，不取第一个。"""
    candidates = [
        TableCandidate(
            name="余额表A",
            headers=["公司", "供应商", "科目", "期初余额", "期末余额"],
            row_count=5,
        ),
        TableCandidate(
            name="余额表B",
            headers=["单位", "往来方", "科目名称", "年初", "年末"],
            row_count=5,
        ),
    ]
    scores = classify_tables(candidates)
    balance_scores = [s for s in scores if s.role == "balance"]
    assert len(balance_scores) == 2


def test_classify_chinese_sheet_name_recognized() -> None:
    """中文 Sheet 名（如"数据表"）通过表头识别，不依赖 Sheet 名。"""
    candidates = [
        TableCandidate(
            name="数据表",
            headers=["日期", "凭证类型", "凭证号", "摘要", "科目", "借方金额", "贷方金额"],
            row_count=200,
        ),
    ]
    scores = classify_tables(candidates)
    # 通过列头关键词识别为序时账角色
    assert any(s.role == "journal" and s.confidence >= 0.80 for s in scores)


# ──────────────────────────────────────────────
# confirm_mapping
# ──────────────────────────────────────────────

def test_confirm_mapping_exact_fields() -> None:
    """确认映射：所有字段列名精确匹配。"""
    candidate = TableCandidate(
        name="序时账",
        headers=["日期", "凭证类型", "凭证号", "摘要", "科目", "借方", "贷方", "辅助核算"],
        row_count=100,
    )
    mapping = ColumnMapping(
        date_col="日期",
        voucher_type_col="凭证类型",
        voucher_no_col="凭证号",
        summary_col="摘要",
        account_col="科目",
        debit_col="借方",
        credit_col="贷方",
        counterparty_aux_col="辅助核算",
    )
    confirmed = confirm_mapping(candidate, mapping)
    assert confirmed.status == DecisionStatus.CONFIRMED


def test_confirm_mapping_missing_optional() -> None:
    """可选字段（辅助核算）缺失不影响确认。"""
    candidate = TableCandidate(
        name="序时账",
        headers=["日期", "凭证号", "摘要", "科目", "借方", "贷方"],
        row_count=50,
    )
    mapping = ColumnMapping(
        date_col="日期",
        voucher_no_col="凭证号",
        summary_col="摘要",
        account_col="科目",
        debit_col="借方",
        credit_col="贷方",
        # voucher_type 和 counterparty_aux 为 None
        counterparty_aux_col=None,
    )
    confirmed = confirm_mapping(candidate, mapping)
    # 必需字段都有映射，可以确认
    assert confirmed.status == DecisionStatus.CONFIRMED


def test_confirm_mapping_missing_required_field() -> None:
    """缺少必需字段（如借方金额）返回 INFERRED。"""
    candidate = TableCandidate(
        name="序时账",
        headers=["日期", "凭证号", "摘要", "科目"],
        row_count=50,
    )
    mapping = ColumnMapping(
        date_col="日期",
        voucher_no_col="凭证号",
        summary_col="摘要",
        account_col="科目",
        # debit 和 credit 未映射
    )
    confirmed = confirm_mapping(candidate, mapping)
    assert confirmed.status != DecisionStatus.CONFIRMED


# ──────────────────────────────────────────────
# build_question_queue
# ──────────────────────────────────────────────

def test_queue_single_question() -> None:
    """队列构建后首问是期间问题。"""
    candidates = [
        TableCandidate(
            name="余额表",
            headers=["公司", "供应商", "科目", "期末余额"],
            row_count=5,
        ),
    ]
    scores = classify_tables(candidates)
    answers = Answers()

    queue = build_question_queue(scores, answers)
    assert queue.has_questions()
    next_q = queue.next()
    assert next_q is not None
    # 第一个问题应该是期间
    assert next_q.category == "period"
    # 队列仍有更多问题（范围、方向、规则）
    assert queue.has_questions()


def test_queue_period_before_columns() -> None:
    """期间问题优先于列映射问题。"""
    candidates = [
        TableCandidate(
            name="序时账",
            headers=["日期", "凭证号", "借方", "贷方", "科目", "摘要"],
            row_count=100,
        ),
    ]
    scores = classify_tables(candidates)
    answers = Answers()

    queue = build_question_queue(scores, answers)
    first = queue.next()
    assert first is not None
    # 第一个问题应该是期间/公司范围
    assert any(kw in first.text.lower() for kw in ["期间", "年度", "period", "范围"])
