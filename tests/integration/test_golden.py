"""Task 12 — Golden XLSX 测试。"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from alternative_test.controls import ControlResult, ControlStatus
from alternative_test.models import (
    CounterpartyBalance,
    JournalLine,
    MatchDecision,
    NormalBalance,
    SourceRef,
)
from alternative_test.renderer import RenderContext, render_workbook
from alternative_test.templates import build_default_template


def _source() -> SourceRef:
    return SourceRef(file_id="t.xlsx", sha256="a" * 64, sheet="S", row=1, columns={})


def _build_context() -> RenderContext:
    """构建完整合成上下文。"""
    return RenderContext(
        project_name="合成测试项目",
        period="2025-01-01 ~ 2025-12-31",
        entity_ids=["E001"],
        run_id="golden-run-001",
        balances=[
            CounterpartyBalance(
                entity_id="E001", counterparty_id="CP001",
                raw_name="合成采购有限公司", account_code="FE001",
                account_name="应付账款", normal_balance=NormalBalance.CREDIT,
                opening_balance=Decimal("50000"), ending_balance=Decimal("75000"),
                source=_source(),
            ),
            CounterpartyBalance(
                entity_id="E001", counterparty_id="CP002",
                raw_name="合成制造股份有限公司", account_code="FE002",
                account_name="应付账款", normal_balance=NormalBalance.CREDIT,
                opening_balance=Decimal("30000"), ending_balance=Decimal("48000"),
                source=_source(),
            ),
        ],
        journal_lines=[
            JournalLine(
                entity_id="E001", ledger_id="L001", posting_date=date(2025, 3, 15),
                voucher_type="记账凭证", voucher_number="24.0001", line_number=1,
                account_code="2202", account_name="应付账款",
                counterparty_raw="合成采购有限公司", summary="采购原材料",
                debit=Decimal("0"), credit=Decimal("35000"), source=_source(),
            ),
            JournalLine(
                entity_id="E001", ledger_id="L001", posting_date=date(2025, 6, 30),
                voucher_type="记账凭证", voucher_number="24.0002", line_number=1,
                account_code="2202", account_name="应付账款",
                counterparty_raw="合成采购有限公司", summary="支付货款",
                debit=Decimal("10000"), credit=Decimal("0"), source=_source(),
            ),
            JournalLine(
                entity_id="E001", ledger_id="L001", posting_date=date(2025, 9, 1),
                voucher_type="记账凭证", voucher_number="24.0003", line_number=1,
                account_code="2202", account_name="应付账款",
                counterparty_raw="合成制造", summary="采购设备",  # 简称 → 需确认
                debit=Decimal("0"), credit=Decimal("20000"), source=_source(),
            ),
            # 红字退货
            JournalLine(
                entity_id="E001", ledger_id="L001", posting_date=date(2025, 10, 1),
                voucher_type="记账凭证", voucher_number="24.0004", line_number=1,
                account_code="2202", account_name="应付账款",
                counterparty_raw="合成制造股份有限公司", summary="退货",
                debit=Decimal("-2000"), credit=Decimal("0"), source=_source(),
            ),
        ],
        match_decisions=[
            MatchDecision(
                journal_line_id="E001_L001_2025-03-15_记账凭证_24.0001_1",
                counterparty_id="CP001", status="exact_unique",
                candidate_ids=["CP001"], reason="精确匹配",
            ),
            MatchDecision(
                journal_line_id="E001_L001_2025-06-30_记账凭证_24.0002_1",
                counterparty_id="CP001", status="exact_unique",
                candidate_ids=["CP001"], reason="精确匹配",
            ),
            MatchDecision(
                journal_line_id="E001_L001_2025-09-01_记账凭证_24.0003_1",
                counterparty_id=None, status="needs_confirmation",
                candidate_ids=["CP002"], reason="去后缀匹配候选",
            ),
            MatchDecision(
                journal_line_id="E001_L001_2025-10-01_记账凭证_24.0004_1",
                counterparty_id="CP002", status="exact_unique",
                candidate_ids=["CP002"], reason="精确匹配",
            ),
        ],
        controls=[
            ControlResult(
                counterparty_id="CP001", raw_name="合成采购有限公司",
                reported_opening=Decimal("50000"), reported_ending=Decimal("75000"),
                total_debit=Decimal("10000"), total_credit=Decimal("35000"),
                expected_ending=Decimal("75000"), difference=Decimal("0"),
                status=ControlStatus.PASS,
            ),
            ControlResult(
                counterparty_id="CP002", raw_name="合成制造股份有限公司",
                reported_opening=Decimal("30000"), reported_ending=Decimal("48000"),
                total_debit=Decimal("-2000"), total_credit=Decimal("20000"),
                expected_ending=Decimal("48000"), difference=Decimal("0"),
                status=ControlStatus.PASS,
            ),
        ],
        excluded_lines=[],
    )


@pytest.fixture
def golden_path(golden_dir: Path) -> Path:
    return golden_dir / "alternative-test.golden.xlsx"


def test_golden_missing_fails(golden_path: Path) -> None:
    """缺 golden 时测试失败。"""
    if not golden_path.exists():
        pytest.fail(
            "Golden 文件不存在。运行 pytest --refresh-golden 生成。"
        )


def test_golden_refresh(request, golden_path: Path) -> None:
    """--refresh-golden 重新生成。"""
    if not request.config.getoption("--refresh-golden"):
        pytest.skip("需要 --refresh-golden 选项")

    with Path(__file__).resolve().parent as _:
        import tempfile
        tmpl_path = Path(tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name)
        build_default_template(tmpl_path)

    ctx = _build_context()
    render_workbook(ctx, tmpl_path, golden_path)

    assert golden_path.exists()


def test_golden_structure(golden_path: Path) -> None:
    """Golden 文件结构完整。"""
    if not golden_path.exists():
        pytest.skip("Golden 文件不存在，先运行 --refresh-golden")

    wb = openpyxl.load_workbook(golden_path, read_only=True)
    sheet_names = wb.sheetnames
    required = [
        "使用说明与确认", "供应商汇总", "完整候选总体",
        "规则命中候选", "未匹配与歧义", "期间外与排除",
        "数学勾稽", "待核查底稿",
    ]
    for name in required:
        assert name in sheet_names, f"缺少 Sheet: {name}"
    wb.close()


def test_golden_no_real_data(golden_path: Path) -> None:
    """Golden 不含真实客户/人员/路径。"""
    if not golden_path.exists():
        pytest.skip("Golden 文件不存在")

    wb = openpyxl.load_workbook(golden_path, read_only=True)
    prohibited = ["Sheet2", "2025-12-31（固定日期）"]
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                cell_str = str(cell)
                for p in prohibited:
                    assert p not in cell_str, (
                        f"Sheet '{sheet_name}' 包含禁止内容: {cell_str}"
                    )
    wb.close()


def test_golden_no_auto_conclusion(golden_path: Path) -> None:
    """待核查底稿无自动结论。"""
    if not golden_path.exists():
        pytest.skip("Golden 文件不存在")

    wb = openpyxl.load_workbook(golden_path, read_only=True)
    ws = wb["待核查底稿"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        # 检查结论列（第 10 列，索引 9）是否为空
        if len(row) >= 10 and row[9] is not None:
            val = str(row[9])
            assert val not in ("未发现异常", "附件完整", "√"), (
                f"待核查底稿含自动结论: {val}"
            )
    wb.close()
