"""Task 3 — 文件探测与 XLSX/XLS/CSV 适配器测试。"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from alternative_test.adapters import (
    ColumnMapping,
    read_rows,
)
from alternative_test.fileio import calc_sha256, detect_format, probe_file

# ──────────────────────────────────────────────
# Fixtures: 用 factories 生成测试文件
# ──────────────────────────────────────────────

@pytest.fixture
def xlsx_balance_path() -> Path:
    """生成余额表 xlsx 的 fixture。"""
    from tests.factories import make_simple_balance_xlsx

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = Path(tmp.name)
    return make_simple_balance_xlsx(path)


@pytest.fixture
def xlsx_journal_path() -> Path:
    """生成序时账 xlsx 的 fixture。"""
    from tests.factories import make_simple_journal_xlsx

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = Path(tmp.name)
    return make_simple_journal_xlsx(path)


@pytest.fixture
def xls_journal_path() -> Path:
    """生成序时账 xls 的 fixture。"""
    from tests.factories import make_simple_journal_xls

    with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
        path = Path(tmp.name)
    return make_simple_journal_xls(path)


@pytest.fixture
def csv_utf8_path() -> Path:
    """生成 UTF-8 CSV 的 fixture。"""
    from tests.factories import make_simple_journal_csv_utf8

    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
        path = Path(tmp.name)
    return make_simple_journal_csv_utf8(path)


@pytest.fixture
def csv_gbk_path() -> Path:
    """生成 GBK CSV（tab 分隔）的 fixture。"""
    from tests.factories import make_simple_journal_csv_gbk

    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
        path = Path(tmp.name)
    return make_simple_journal_csv_gbk(path)


# ──────────────────────────────────────────────
# SHA-256
# ──────────────────────────────────────────────

def test_calc_sha256_consistent(xlsx_balance_path: Path) -> None:
    """两次计算同一文件 SHA-256 一致。"""
    h1 = calc_sha256(xlsx_balance_path)
    h2 = calc_sha256(xlsx_balance_path)
    assert h1 == h2
    assert len(h1) == 64


def test_calc_sha256_input_preserved(xlsx_balance_path: Path) -> None:
    """计算哈希后文件内容不变。"""
    h1 = calc_sha256(xlsx_balance_path)
    h2 = calc_sha256(xlsx_balance_path)
    assert h1 == h2


# ──────────────────────────────────────────────
# 探测: detect_format
# ──────────────────────────────────────────────

def test_detect_xlsx(xlsx_balance_path: Path) -> None:
    """识别 xlsx 格式。"""
    fmt = detect_format(xlsx_balance_path)
    assert fmt == "xlsx"


def test_detect_xls(xls_journal_path: Path) -> None:
    """识别 xls 格式。"""
    fmt = detect_format(xls_journal_path)
    assert fmt == "xls"


def test_detect_csv_utf8(csv_utf8_path: Path) -> None:
    """识别 csv 格式。"""
    fmt = detect_format(csv_utf8_path)
    assert fmt == "csv"


# ──────────────────────────────────────────────
# 探测: probe_file
# ──────────────────────────────────────────────

def test_probe_xlsx_multi_sheet(xlsx_journal_path: Path) -> None:
    """探测多 Sheet xlsx — 所有 Sheet 列出。"""
    probe = probe_file(xlsx_journal_path)
    assert probe.format == "xlsx"
    assert len(probe.sha256) == 64
    assert len(probe.candidates) >= 2  # Sheet1 + Sheet2


def test_probe_xls(xls_journal_path: Path) -> None:
    """探测 xls 文件正常。"""
    probe = probe_file(xls_journal_path)
    assert probe.format == "xls"
    assert len(probe.candidates) >= 1


def test_probe_csv(csv_utf8_path: Path) -> None:
    """探测 CSV 文件正常。"""
    probe = probe_file(csv_utf8_path)
    assert probe.format == "csv"
    assert len(probe.candidates) >= 1


def test_probe_records_row_count(xlsx_balance_path: Path) -> None:
    """探测记录行数（含表头）。"""
    probe = probe_file(xlsx_balance_path)
    candidate = probe.candidates[0]
    assert candidate.row_count >= 3


# ──────────────────────────────────────────────
# 适配器: read_rows
# ──────────────────────────────────────────────

def test_read_xlsx_with_mapping(xlsx_journal_path: Path) -> None:
    """通过列映射读取 xlsx 行。"""
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
    rows = list(read_rows(xlsx_journal_path, mapping, sheet="Sheet1"))
    assert len(rows) == 2
    assert rows[0]["日期"] is not None


def test_read_xls_with_mapping(xls_journal_path: Path) -> None:
    """通过列映射读取 xls 行。"""
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
    rows = list(read_rows(xls_journal_path, mapping, sheet="序时账"))
    assert len(rows) == 2


def test_read_csv_utf8_with_mapping(csv_utf8_path: Path) -> None:
    """通过列映射读取 CSV（自动检测分隔符和编码）。"""
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
    rows = list(read_rows(csv_utf8_path, mapping))
    assert len(rows) == 2


def test_read_csv_gbk_with_mapping(csv_gbk_path: Path) -> None:
    """通过列映射读取 GBK tab 分隔 CSV。"""
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
    rows = list(read_rows(csv_gbk_path, mapping))
    assert len(rows) == 2


# ──────────────────────────────────────────────
# 公式处理
# ──────────────────────────────────────────────

def test_formula_cell_preserved() -> None:
    """公式单元格不执行，保留公式文本。"""
    from tests.factories import make_xlsx_with_formulas

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = Path(tmp.name)
    make_xlsx_with_formulas(path)

    mapping = ColumnMapping(
        date_col=None,
        voucher_type_col=None,
        voucher_no_col=None,
        summary_col=None,
        account_col=None,
        debit_col="金额",
        credit_col="公式列",
        counterparty_aux_col=None,
    )
    rows = list(read_rows(path, mapping, sheet="数据"))
    assert len(rows) == 2
    # 公式列应为公式文本或 None，不能是计算结果
    formula_val = rows[0].get("公式列")
    assert formula_val is None or isinstance(formula_val, str)


# ──────────────────────────────────────────────
# 伪装扩展名
# ──────────────────────────────────────────────

def test_misnamed_xlsx_as_xls() -> None:
    """xlsx 文件伪装为 .xls 扩展名仍可正确探测。"""
    from tests.factories import make_xlsx

    with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
        path = Path(tmp.name)
    make_xlsx(path, {"数据": [["A", "B"], [1, 2]]})

    probe = probe_file(path)
    # 按签名识别为 xlsx（OOXML），不是 xls（OLE）
    assert probe.format == "xlsx"


# ──────────────────────────────────────────────
# 输入不变性
# ──────────────────────────────────────────────

def test_input_hash_stable_after_read(xlsx_balance_path: Path) -> None:
    """读取后输入文件哈希不变。"""
    h_before = calc_sha256(xlsx_balance_path)
    probe_file(xlsx_balance_path)
    mapping = ColumnMapping(
        date_col=None,
        voucher_type_col=None,
        voucher_no_col=None,
        summary_col=None,
        account_col=None,
        debit_col=None,
        credit_col=None,
        counterparty_aux_col=None,
    )
    list(read_rows(xlsx_balance_path, mapping))
    h_after = calc_sha256(xlsx_balance_path)
    assert h_before == h_after
