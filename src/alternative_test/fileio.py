"""文件探测 — 识别格式、记录表头、计算哈希。"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import xlrd

# ──────────────────────────────────────────────
# 哈希
# ──────────────────────────────────────────────

def calc_sha256(path: Path) -> str:
    """计算文件 SHA-256（读二进制）。"""
    sha = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            sha.update(chunk)
    return sha.hexdigest()


# ──────────────────────────────────────────────
# 数据类
# ──────────────────────────────────────────────

@dataclass
class TableCandidate:
    """工作表/表格候选。"""

    name: str          # Sheet 名或 "csv_default"
    headers: list[str]  # 首行表头
    row_count: int      # 总行数（含表头）
    has_formulas: bool = False
    has_merged_cells: bool = False
    is_hidden: bool = False


@dataclass
class FileProbe:
    """文件探测结果。"""

    path: Path
    format: str        # "xlsx" | "xls" | "csv"
    sha256: str
    size: int
    candidates: list[TableCandidate] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# ──────────────────────────────────────────────
# 格式检测
# ──────────────────────────────────────────────

def detect_format(path: Path) -> str:
    """按文件签名识别格式。"""
    with open(path, "rb") as f:
        header = f.read(8)

    # OOXML (xlsx) 以 PK\x03\x04 开头
    if header[:4] == b"PK\x03\x04":
        return "xlsx"

    # OLE (xls) 以 D0CF11E0 开头
    if header[:4] == b"\xd0\xcf\x11\xe0":
        return "xls"

    # 文本文件
    return "csv"


# ──────────────────────────────────────────────
# 探测
# ──────────────────────────────────────────────

def probe_file(path: Path) -> FileProbe:
    """探测文件：格式、Sheet、表头、行数、SHA-256。"""
    fmt = detect_format(path)
    sha = calc_sha256(path)
    size = path.stat().st_size

    candidates: list[TableCandidate] = []
    warnings: list[str] = []

    if fmt == "xlsx":
        # openpyxl read_only 模式
        # 若扩展名不匹配（如 .xls 伪装），用 BytesIO 绕过扩展名校验
        if path.suffix.lower() in (".xlsx", ".xlsm"):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        else:
            from io import BytesIO
            wb = openpyxl.load_workbook(
                BytesIO(path.read_bytes()), read_only=True, data_only=False
            )
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # 读取首行作为表头
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            headers = [str(c) if c is not None else "" for c in (first_row or [])]

            # 计数行数
            row_count = 0
            has_formulas = False
            for row in ws.iter_rows(values_only=False):
                row_count += 1
                if not has_formulas:
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            has_formulas = True
                            break

            # 检查隐藏和合并单元格（read_only 模式下 merged_cells 不可用）
            try:
                has_merged = bool(ws.merged_cells.ranges)
            except AttributeError:
                has_merged = False
            candidates.append(TableCandidate(
                name=sheet_name,
                headers=headers,
                row_count=row_count,
                has_formulas=has_formulas,
                has_merged_cells=has_merged,
            ))
        wb.close()

    elif fmt == "xls":
        wb = xlrd.open_workbook(str(path), on_demand=True)
        for idx in range(wb.nsheets):
            ws = wb.sheet_by_index(idx)
            sheet_name = ws.name
            headers = [str(ws.cell_value(0, c)) for c in range(ws.ncols)] if ws.nrows > 0 else []
            candidates.append(TableCandidate(
                name=sheet_name,
                headers=headers,
                row_count=ws.nrows,
                has_merged_cells=bool(ws.merged_cells),
            ))
        wb.release_resources()

    else:
        # CSV — 检测分隔符和编码
        encoding, delimiter, headers, row_count = _sniff_csv(path)
        candidates.append(TableCandidate(
            name="csv_default",
            headers=headers,
            row_count=row_count,
        ))

    return FileProbe(
        path=path,
        format=fmt,
        sha256=sha,
        size=size,
        candidates=candidates,
        warnings=warnings,
    )


def _sniff_csv(path: Path) -> tuple[str, str, list[str], int]:
    """检测 CSV 编码、分隔符和表头。"""
    # 先尝试常见编码
    for encoding in ["utf-8", "utf-8-sig", "gb18030", "gbk", "gb2312"]:
        try:
            with open(path, encoding=encoding) as f:
                lines = f.readlines()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        # 兜底
        encoding = "utf-8"
        with open(path, encoding=encoding, errors="replace") as f:
            lines = f.readlines()

    # 检测分隔符：统计 tab vs 逗号
    sample = "".join(lines[:min(len(lines), 20)])
    tabs = sample.count("\t")
    commas = sample.count(",")
    delimiter = "\t" if tabs > commas else ","

    # 解析表头
    if lines:
        headers = [h.strip().strip('"') for h in lines[0].rstrip("\r\n").split(delimiter)]
    else:
        headers = []
    row_count = len(lines)

    return encoding, delimiter, headers, row_count
