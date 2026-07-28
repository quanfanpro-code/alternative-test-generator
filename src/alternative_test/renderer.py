"""安全渲染 — 汇总表+N明细Sheet，零合并单元格，待核查字段全部空。"""

from __future__ import annotations

import copy
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .errors import OutputConflictError
from .models import CounterpartyBalance, JournalLine, MatchDecision, ProcedureResult
from .controls import ControlResult
from .selection import SelectionResult
from .population import PopulationResult


# ── 样式 ──

DATA_FONT = Font(name="宋体", size=10)
HEADING_FONT = Font(name="宋体", size=10, bold=True)
AMOUNT_FONT = Font(name="Arial Narrow", size=10)
AMOUNT_FMT = "#,##0.00"
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


# ── 上下文 ──

@dataclass
class RenderContext:
    project_name: str = ""
    period: str = ""
    entity_ids: list[str] = field(default_factory=list)
    side: str = "liability"  # "liability" | "asset"
    balances: list[CounterpartyBalance] = field(default_factory=list)
    journal_lines: list[JournalLine] = field(default_factory=list)
    match_decisions: list[MatchDecision] = field(default_factory=list)
    controls: list[ControlResult] = field(default_factory=list)
    selection: SelectionResult | None = None
    procedure_results: dict[str, ProcedureResult] = field(default_factory=dict)
    excluded_lines: list[tuple[JournalLine, str]] = field(default_factory=list)
    population: PopulationResult | None = None
    run_id: str = ""
    input_hashes: dict[str, str] = field(default_factory=dict)
    check_overrides: dict[str, list[dict]] = field(default_factory=dict)


def render_workbook(context: RenderContext, template: Path, output: Path) -> Path:
    """渲染工作簿 — 临时保存、验证后原子改名。"""
    if output.exists():
        raise OutputConflictError(f"输出文件已存在: {output}")

    tmp_dir = Path(tempfile.mkdtemp(prefix="alt_test_"))
    tmp_output = tmp_dir / output.name
    try:
        _do_render(context, template, tmp_output)
        wb = openpyxl.load_workbook(tmp_output, read_only=True)
        sheet_count = len(wb.sheetnames)
        wb.close()
        if sheet_count < 2:
            raise RuntimeError(f"渲染失败：仅 {sheet_count} 个工作表")
        shutil.move(str(tmp_output), str(output))
        return output
    finally:
        if tmp_dir.exists():
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _do_render(context: RenderContext, template: Path, output: Path) -> None:
    """实际渲染逻辑。"""
    wb = openpyxl.load_workbook(template)

    # 加载合同
    from .templates import TemplateContract
    contract = TemplateContract.from_yaml(
        Path(__file__).resolve().parent.parent.parent / "assets" / "templates" / "template-contract.yaml"
    )
    tdef = contract.templates[context.side]

    # 1. 填汇总表
    _fill_summary(wb[tdef.summary_sheet], context)

    # 2. 为每个对手方新建 Sheet，从模板复制结构
    template_sheet = wb[tdef.template_sheet]
    for balance in context.balances:
        short_name = balance.raw_name[:20] if balance.raw_name else "未知"
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = short_name
        _fill_detail_sheet(new_sheet, balance, context, tdef)

    wb.save(output)
    wb.close()


def _fill_summary(ws, ctx: RenderContext) -> None:
    """填充汇总表数据。"""
    # 找数据起始行：跳过标题和信息行
    data_start = 12  # 汇总表表头固定在第12行附近
    r = data_start + 1

    total_original = Decimal("0")
    total_confirmed = Decimal("0")

    for i, b in enumerate(ctx.balances, 1):
        ctrl = next((c for c in ctx.controls if c.counterparty_id == b.counterparty_id), None)
        confirmed = ctrl.reported_ending if ctrl and ctrl.status == "PASS" else Decimal("0")
        total_original += b.ending_balance
        total_confirmed += confirmed
        row_data = [
            b.account_code or f"{ctx.project_name}-{i:03d}",
            b.raw_name,
            "人民币",
            b.ending_balance,
            confirmed,
            "",
        ]
        for ci, val in enumerate(row_data):
            cell = ws.cell(row=r, column=2 + ci)
            if isinstance(val, Decimal):
                cell.value = float(val)
                cell.number_format = AMOUNT_FMT
                cell.font = AMOUNT_FONT
            else:
                cell.value = val
                cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = CENTER
        r += 1

    # 合计行（替换"共 0 行"）
    # 找到"共 N 行"的位置并更新
    for row in ws.iter_rows(min_row=data_start, max_row=data_start + 5):
        for cell in row:
            if cell.value and "共" in str(cell.value) and "行" in str(cell.value):
                cell.value = f"共 {len(ctx.balances)} 行"
                cell_credit = ws.cell(row=cell.row, column=6)
                cell_credit.value = float(total_original)
                cell_credit.font = AMOUNT_FONT
                cell_credit.number_format = AMOUNT_FMT
                cell_confirmed = ws.cell(row=cell.row, column=7)
                cell_confirmed.value = float(total_confirmed)
                cell_confirmed.font = AMOUNT_FONT
                cell_confirmed.number_format = AMOUNT_FMT


def _fill_detail_sheet(ws, balance: CounterpartyBalance, ctx: RenderContext, tdef) -> None:
    """填充单张明细 Sheet。"""
    # 找三个板块的数据起始位置
    section_positions = _locate_sections(ws)
    lines = _get_matched_lines(balance, ctx)
    post_lines = _get_post_period_lines(balance, ctx)

    # 负债方：section 0=贷方(增加/credit), 1=借方(减少/debit), 2=期后付款(debit)
    # 资产方：section 0=借方(增加/debit), 1=贷方(减少/credit), 2=期后收款(credit)
    if ctx.side == "liability":
        sec_increase = (section_positions[0], "credit")
        sec_decrease = (section_positions[1], "debit")
        sec_post = (section_positions[2], "debit")
    else:
        sec_increase = (section_positions[0], "debit")
        sec_decrease = (section_positions[1], "credit")
        sec_post = (section_positions[2], "credit")

    for (sec_row, _), direction in [sec_increase, sec_decrease]:
        filtered = [l for l in lines if
                    (direction == "credit" and l.credit != Decimal("0")) or
                    (direction == "debit" and l.debit != Decimal("0"))]
        _fill_section_data(ws, sec_row + 3, filtered, direction)

    # 期后
    if post_lines:
        post_filtered = [l for l in post_lines if sec_post[1] == "debit" and l.debit != Decimal("0") or
                         sec_post[1] == "credit" and l.credit != Decimal("0")]
        _fill_section_data(ws, sec_post[0] + 3, post_filtered, sec_post[1])


def _locate_sections(ws) -> list[tuple[int, str]]:
    """定位三个板块起始行。"""
    positions = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                continue
            s = str(val)
            if "一、" in s or "贷方发生额" in s or "借方发生额" in s:
                if ("一、" in s or (not positions)) and not any(p[0] == r for p in positions):
                    positions.append((r, s))
                    break
            elif "二、" in s:
                if not any(p[0] == r for p in positions):
                    positions.append((r, s))
                    break
            elif "三、" in s or "期后" in s:
                if not any(p[0] == r for p in positions):
                    positions.append((r, s))
                    break
        if len(positions) >= 3:
            break
    return positions if len(positions) == 3 else [(1, ""), (1, ""), (1, "")]


def _get_matched_lines(balance: CounterpartyBalance, ctx: RenderContext) -> list[JournalLine]:
    """获取该对手方的期间内行。"""
    cp_id = balance.counterparty_id
    matched_ids = {d.journal_line_id for d in ctx.match_decisions
                   if d.counterparty_id == cp_id and d.status in ("exact_unique", "confirmed")}
    return [l for l in ctx.journal_lines
            if _line_id(l) in matched_ids]


def _get_post_period_lines(balance: CounterpartyBalance, ctx: RenderContext) -> list[JournalLine]:
    """获取期间外行（期后）。"""
    cp_id = balance.counterparty_id
    return [l for l in ctx.excluded_lines or []
            if hasattr(l, '__getitem__') and l[0].entity_id == balance.entity_id
            for mid in [d for d in ctx.match_decisions if d.counterparty_id == cp_id]
            if _line_id(l[0]) == mid.journal_line_id]


def _fill_section_data(ws, start_row: int, lines: list[JournalLine], direction: str) -> None:
    """在板块中填充数据行。"""
    if not lines:
        return
    r = start_row
    total = Decimal("0")
    for i, line in enumerate(lines, 1):
        amount = line.credit if direction == "credit" else line.debit
        row_data = [
            i,
            line.posting_date,
            line.voucher_type,
            line.voucher_number,
            line.summary,
            line.account_name,
            amount,
        ]
        for ci, val in enumerate(row_data):
            cell = ws.cell(row=r, column=2 + ci)
            if isinstance(val, Decimal):
                cell.value = float(val)
                cell.number_format = AMOUNT_FMT
                cell.font = AMOUNT_FONT
            elif isinstance(val, date):
                cell.value = val
                cell.number_format = "YYYY-MM-DD"
                cell.font = DATA_FONT
            else:
                cell.value = val
                cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = CENTER if ci in (0, 1, 2, 3) else LEFT
        total += amount
        r += 1

    # 更新小计行
    for row in ws.iter_rows(min_row=r, max_row=r + 2):
        for cell in row:
            if cell.value and str(cell.value) == "0":
                cell.value = len(lines)
            elif cell.value == "0" and cell.column >= 3:
                pass  # 不覆盖其他0值


def _line_id(line: JournalLine) -> str:
    return f"JL{line.line_number}"
