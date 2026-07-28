"""渲染器测试 — 新格式：汇总表+N明细Sheet，按负债/资产方输出。"""

from __future__ import annotations

import tempfile
from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from alternative_test.errors import OutputConflictError
import shutil

from alternative_test.renderer import RenderContext, render_workbook
from alternative_test.models import (
    CounterpartyBalance,
    JournalLine,
    MatchDecision,
    NormalBalance,
    SourceRef,
)
from alternative_test.controls import ControlResult, ControlStatus
from alternative_test.population import PopulationResult
from alternative_test.selection import SelectionResult


def _src() -> SourceRef:
    return SourceRef(file_id="t.xlsx", sha256="a" * 64, sheet="S", row=1, columns={})


ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets" / "templates"


@pytest.fixture
def liability_tmpl() -> Path:
    src = ASSETS_DIR / "替代测试-负债方-v1.xlsx"
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        dst = Path(tmp.name)
    if src.exists():
        shutil.copy2(src, dst)
    else:
        pytest.skip("预构建模板不存在，先运行 build_default_template.py")
    return dst


@pytest.fixture
def asset_tmpl() -> Path:
    src = ASSETS_DIR / "替代测试-资产方-v1.xlsx"
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        dst = Path(tmp.name)
    if src.exists():
        shutil.copy2(src, dst)
    else:
        pytest.skip("预构建模板不存在，先运行 build_default_template.py")
    return dst


@pytest.fixture
def liability_ctx() -> RenderContext:
    ctx = RenderContext(
        project_name="应付账款-替代测试",
        period="2025-01-01 ~ 2025-12-31",
        entity_ids=["E001"], run_id="test-001", side="liability",
        balances=[
            CounterpartyBalance(entity_id="E001", counterparty_id="CP001",
                raw_name="合成采购有限公司", account_code="FE001",
                account_name="应付账款", normal_balance=NormalBalance.CREDIT,
                opening_balance=Decimal("10000"), ending_balance=Decimal("25000"),
                source=_src()),
        ],
        journal_lines=[
            JournalLine(entity_id="E001", ledger_id="L001",
                posting_date=date(2025, 6, 30), voucher_type="记账凭证",
                voucher_number="24.0001", line_number=1,
                account_code="2202", account_name="应付账款",
                counterparty_raw="合成采购有限公司", summary="采购原材料",
                debit=Decimal("0"), credit=Decimal("15000"), source=_src()),
            JournalLine(entity_id="E001", ledger_id="L001",
                posting_date=date(2025, 9, 15), voucher_type="记账凭证",
                voucher_number="24.0002", line_number=1,
                account_code="2202", account_name="应付账款",
                counterparty_raw="合成采购有限公司", summary="付款",
                debit=Decimal("2000"), credit=Decimal("0"), source=_src()),
        ],
        match_decisions=[
            MatchDecision(journal_line_id="JL001", counterparty_id="CP001",
                status="exact_unique", candidate_ids=["CP001"], reason="精确匹配"),
            MatchDecision(journal_line_id="JL002", counterparty_id="CP001",
                status="exact_unique", candidate_ids=["CP001"], reason="精确匹配"),
        ],
        controls=[
            ControlResult(counterparty_id="CP001", raw_name="合成采购有限公司",
                reported_opening=Decimal("10000"), reported_ending=Decimal("25000"),
                total_debit=Decimal("2000"), total_credit=Decimal("15000"),
                expected_ending=Decimal("23000"), difference=Decimal("2000"),
                status=ControlStatus.DIFFERENCE),
        ],
        selection=SelectionResult(population=[]),
    )
    return ctx


@pytest.fixture
def asset_ctx() -> RenderContext:
    ctx = RenderContext(
        project_name="应收账款-替代测试", period="2025-01-01 ~ 2025-12-31",
        entity_ids=["E001"], run_id="test-002", side="asset",
        balances=[
            CounterpartyBalance(entity_id="E001", counterparty_id="CP001",
                raw_name="合成销售有限公司", account_code="AR001",
                account_name="应收账款", normal_balance=NormalBalance.DEBIT,
                opening_balance=Decimal("30000"), ending_balance=Decimal("50000"),
                source=_src()),
        ],
        journal_lines=[
            JournalLine(entity_id="E001", ledger_id="L001",
                posting_date=date(2025, 3, 1), voucher_type="记账凭证",
                voucher_number="24.0001", line_number=1,
                account_code="1122", account_name="应收账款",
                counterparty_raw="合成销售有限公司", summary="销售",
                debit=Decimal("20000"), credit=Decimal("0"), source=_src()),
        ],
        match_decisions=[
            MatchDecision(journal_line_id="JL001", counterparty_id="CP001",
                status="exact_unique", candidate_ids=["CP001"], reason="精确匹配"),
        ],
        controls=[
            ControlResult(counterparty_id="CP001", raw_name="合成销售有限公司",
                reported_opening=Decimal("30000"), reported_ending=Decimal("50000"),
                total_debit=Decimal("20000"), total_credit=Decimal("0"),
                expected_ending=Decimal("50000"), difference=Decimal("0"),
                status=ControlStatus.PASS),
        ],
        selection=SelectionResult(population=[]),
    )
    return ctx


# ── 负债方渲染 ──

def test_liability_render_creates_output(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """渲染负债方上下文生成文件。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    result = render_workbook(liability_ctx, liability_tmpl, o)
    assert result.exists()


def test_liability_has_summary_sheet(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """输出含"替代结果汇总表"。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    render_workbook(liability_ctx, liability_tmpl, o)
    wb = openpyxl.load_workbook(o, read_only=True)
    assert "替代结果汇总表" in wb.sheetnames
    wb.close()


def test_liability_each_counterparty_own_sheet(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """每个供应商生成独立 Sheet。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    render_workbook(liability_ctx, liability_tmpl, o)
    wb = openpyxl.load_workbook(o, read_only=True)
    # 应该有汇总表 + 每个供应商一个 Sheet
    supplier_sheets = [s for s in wb.sheetnames if s not in ("替代结果汇总表", "替代测试表模板")]
    assert len(supplier_sheets) >= 1
    wb.close()


def test_liability_data_in_correct_section(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """贷方发生额数据填入贷方板块。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    render_workbook(liability_ctx, liability_tmpl, o)
    wb = openpyxl.load_workbook(o, read_only=True)
    sup_sheets = [s for s in wb.sheetnames if s != "替代结果汇总表"]
    ws = wb[sup_sheets[0]]
    found_credit = False
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell == "贷方发生额" or (isinstance(cell, str) and "贷方发生额" in cell):
                found_credit = True
    assert found_credit
    wb.close()


def test_liability_check_columns_empty(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """检查列（入库单日期/数量等）生成后为空。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    render_workbook(liability_ctx, liability_tmpl, o)
    wb = openpyxl.load_workbook(o, read_only=True)
    sup_sheets = [s for s in wb.sheetnames if s != "替代结果汇总表"]
    ws = wb[sup_sheets[0]]
    # 数据行之后检查列 J-M 应为空（或含表头文字"日期""数量"）
    has_data = False
    for r_idx, row in enumerate(ws.iter_rows(min_row=12, values_only=True)):
        if row[7] is not None and row[7] != "" and row[7] != "金    额":
            has_data = True
        if has_data and r_idx > 0:
            for c in [9, 10, 11]:  # 检查列
                if c < len(row):
                    assert row[c] is None or row[c] == "", (
                        f"检查列{c}不应有预填数据: {row[c]}"
                    )
    wb.close()


# ── 资产方渲染 ──

def test_asset_render_creates_output(asset_tmpl: Path, asset_ctx: RenderContext) -> None:
    """渲染资产方上下文生成文件。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    result = render_workbook(asset_ctx, asset_tmpl, o)
    assert result.exists()


def test_asset_sections_correct(asset_tmpl: Path, asset_ctx: RenderContext) -> None:
    """资产方模板含借方发生额、贷方发生额、期后收款。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    render_workbook(asset_ctx, asset_tmpl, o)
    wb = openpyxl.load_workbook(o, read_only=True)
    sup_sheets = [s for s in wb.sheetnames if s != "替代结果汇总表"]
    ws = wb[sup_sheets[0]]
    found_debit = found_credit = found_post = False
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            s = str(cell) if cell else ""
            if "借方发生额" in s:
                found_debit = True
            if "贷方发生额" in s:
                found_credit = True
            if "期后收款" in s:
                found_post = True
    assert found_debit
    assert found_credit
    assert found_post
    wb.close()


# ── 通用 ──

def test_render_output_conflict(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """输出已存在时抛错误。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    with pytest.raises(OutputConflictError):
        render_workbook(liability_ctx, liability_tmpl, o)


def test_render_no_auto_conclusion(liability_tmpl: Path, liability_ctx: RenderContext) -> None:
    """输出不含√、未发现异常、附件完整。"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        o = Path(tmp.name)
    o.unlink(missing_ok=True)
    render_workbook(liability_ctx, liability_tmpl, o)
    wb = openpyxl.load_workbook(o, read_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                s = str(cell)
                for w in ["√", "未发现异常", "附件完整"]:
                    assert w not in s, f"Sheet '{name}' 含自动结论: {s}"
    wb.close()
