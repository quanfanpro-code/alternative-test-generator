"""测试工厂 — 生成合成 xlsx/xls/csv 文件。"""

from __future__ import annotations

import csv
import hashlib
from pathlib import Path
from typing import Any

import openpyxl


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ──────────────────────────────────────────────
# XLSX
# ──────────────────────────────────────────────

def make_xlsx(
    path: Path,
    sheets: dict[str, list[list[Any]]],
) -> Path:
    """生成 xlsx 文件。key=Sheet名, value=行列表（首行为表头）。"""
    wb = openpyxl.Workbook()
    # 删除默认 sheet
    wb.remove(wb.active)
    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(title=sheet_name)
        for row_data in rows:
            ws.append(row_data)
    wb.save(path)
    wb.close()
    return path


def make_simple_balance_xlsx(path: Path) -> Path:
    """生成简单的余额表 xlsx。"""
    return make_xlsx(
        path,
        {
            "余额表": [
                ["公司名称", "供应商名称", "科目代码", "科目名称", "期初余额", "期末余额"],
                ["测试公司A", "合成采购有限公司", "FE001", "应付账款", 50000, 80000],
                ["测试公司A", "合成制造有限公司", "FE002", "应付账款", 30000, 45000],
            ],
        },
    )


def make_simple_journal_xlsx(path: Path) -> Path:
    """生成简单的序时账 xlsx（多 sheet）。"""
    return make_xlsx(
        path,
        {
            "Sheet1": [
                ["日期", "凭证类型", "凭证号", "摘要", "科目", "借方", "贷方", "辅助核算"],
                ["2025-06-01", "记账凭证", "24.0001", "采购原材料", "应付账款", 10000, 0, "合成采购有限公司"],
                ["2025-06-15", "记账凭证", "24.0002", "支付货款", "应付账款", 0, 8000, "合成采购有限公司"],
            ],
            "Sheet2": [
                ["其他数据", "备注"],
                ["无关", "忽略"],
            ],
        },
    )


def make_xlsx_with_formulas(path: Path) -> Path:
    """生成含公式单元格的 xlsx。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="数据")
    ws.append(["编号", "金额", "公式列"])
    ws.append([1, 100, "=B2*2"])
    ws.append([2, 200, "=B3*2"])
    wb.save(path)
    wb.close()
    return path


# ──────────────────────────────────────────────
# XLS
# ──────────────────────────────────────────────

def make_xls(path: Path, sheets: dict[str, list[list[Any]]]) -> Path:
    """生成 xls 文件。使用 openpyxl 写为 xls 兼容格式（实际写 xlsx 改扩展名不可行）。
    改用 xlwt 的替代：实际用 openpyxl 写 OOXML 格式但包装成 xls 测试 xlrd 读取。
    此函数写真实 .xls 格式。
    """
    import xlwt

    wb = xlwt.Workbook()
    for sheet_name, rows in sheets.items():
        ws = wb.add_sheet(sheet_name)
        for r, row_data in enumerate(rows):
            for c, val in enumerate(row_data):
                ws.write(r, c, val)
    wb.save(str(path))
    return path


def make_simple_journal_xls(path: Path) -> Path:
    """生成简单序时账 xls。"""
    return make_xls(
        path,
        {
            "序时账": [
                ["日期", "凭证类型", "凭证号", "摘要", "科目", "借方", "贷方", "辅助核算"],
                ["2025-06-01", "记账凭证", "24.0001", "采购原材料", "应付账款", 10000, 0, "合成采购有限公司"],
                ["2025-06-15", "记账凭证", "24.0002", "支付货款", "应付账款", 0, 8000, "合成采购有限公司"],
            ],
        },
    )


# ──────────────────────────────────────────────
# CSV
# ──────────────────────────────────────────────

def make_csv(
    path: Path,
    rows: list[list[Any]],
    encoding: str = "utf-8",
    delimiter: str = ",",
) -> Path:
    """生成 CSV 文件。"""
    with open(path, "w", newline="", encoding=encoding) as f:
        writer = csv.writer(f, delimiter=delimiter)
        for row in rows:
            writer.writerow(row)
    return path


def make_simple_journal_csv_utf8(path: Path) -> Path:
    """生成 UTF-8 CSV 序时账。"""
    return make_csv(
        path,
        [
            ["日期", "凭证类型", "凭证号", "摘要", "科目", "借方", "贷方", "辅助核算"],
            ["2025-06-01", "记账凭证", "24.0001", "采购原材料", "应付账款", "10000", "0", "合成采购有限公司"],
            ["2025-06-15", "记账凭证", "24.0002", "支付货款", "应付账款", "0", "8000", "合成采购有限公司"],
        ],
        encoding="utf-8",
    )


def make_simple_journal_csv_gbk(path: Path) -> Path:
    """生成 GBK CSV 序时账（tab 分隔）。"""
    return make_csv(
        path,
        [
            ["日期", "凭证类型", "凭证号", "摘要", "科目", "借方", "贷方", "辅助核算"],
            ["2025-06-01", "记账凭证", "24.0001", "采购原材料", "应付账款", "10000", "0", "合成采购有限公司"],
            ["2025-06-15", "记账凭证", "24.0002", "支付货款", "应付账款", "0", "8000", "合成采购有限公司"],
        ],
        encoding="gb18030",
        delimiter="\t",
    )
