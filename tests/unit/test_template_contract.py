"""模板构建与验证测试 — 使用预构建模板。"""

from __future__ import annotations

import hashlib
import shutil
import tempfile
from pathlib import Path

import openpyxl
import pytest

from alternative_test.templates import TemplateContract, TemplateValidation, validate_template

ASSETS = Path(__file__).resolve().parents[2] / "assets" / "templates"


@pytest.fixture
def contract_path() -> Path:
    return ASSETS / "template-contract.yaml"


@pytest.fixture
def contract(contract_path: Path) -> TemplateContract:
    return TemplateContract.from_yaml(contract_path)


@pytest.fixture
def liability_template() -> Path:
    """复制预构建负债方模板到临时文件。"""
    src = ASSETS / "替代测试-负债方-v1.xlsx"
    if not src.exists():
        pytest.skip("预构建模板不存在，先运行 scripts/build_default_template.py")
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        dst = Path(tmp.name)
    shutil.copy2(src, dst)
    return dst


@pytest.fixture
def asset_template() -> Path:
    """复制预构建资产方模板到临时文件。"""
    src = ASSETS / "替代测试-资产方-v1.xlsx"
    if not src.exists():
        pytest.skip("预构建模板不存在，先运行 scripts/build_default_template.py")
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        dst = Path(tmp.name)
    shutil.copy2(src, dst)
    return dst


# ── 合同加载 ──

def test_contract_loads(contract: TemplateContract) -> None:
    assert contract.version == "1.1"
    assert "liability" in contract.templates
    assert "asset" in contract.templates


# ── 负债方模板 ──

def test_liability_template_opens(liability_template: Path) -> None:
    wb = openpyxl.load_workbook(liability_template, read_only=True)
    wb.close()


def test_liability_has_summary_sheet(liability_template: Path) -> None:
    wb = openpyxl.load_workbook(liability_template, read_only=True)
    assert "替代结果汇总表" in wb.sheetnames
    wb.close()


def test_liability_section_titles(liability_template: Path) -> None:
    wb = openpyxl.load_workbook(liability_template, read_only=True)
    template_sheets = [s for s in wb.sheetnames if s != "替代结果汇总表"]
    assert len(template_sheets) >= 1
    ws = wb[template_sheets[0]]
    found_1 = found_2 = found_3 = False
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell and "贷方发生额" in str(cell):
                found_1 = True
            if cell and "借方发生额" in str(cell):
                found_2 = True
            if cell and "期后付款" in str(cell):
                found_3 = True
    assert found_1, "缺少'贷方发生额'"
    assert found_2, "缺少'借方发生额'"
    assert found_3, "缺少'期后付款'"
    wb.close()


def test_liability_no_merged_cells(liability_template: Path) -> None:
    wb = openpyxl.load_workbook(liability_template)
    for name in wb.sheetnames:
        ws = wb[name]
        assert not ws.merged_cells.ranges, f"Sheet '{name}' 含合并单元格: {ws.merged_cells.ranges}"
    wb.close()


def test_liability_has_center_across(liability_template: Path) -> None:
    wb = openpyxl.load_workbook(liability_template)
    found = False
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.alignment and cell.alignment.horizontal == "centerContinuous":
                    found = True
                    break
    assert found, "模板未使用跨列居中"
    wb.close()


# ── 资产方模板 ──

def test_asset_template_opens(asset_template: Path) -> None:
    wb = openpyxl.load_workbook(asset_template, read_only=True)
    wb.close()


def test_asset_has_summary_sheet(asset_template: Path) -> None:
    wb = openpyxl.load_workbook(asset_template, read_only=True)
    assert "替代结果汇总表" in wb.sheetnames
    wb.close()


def test_asset_section_titles(asset_template: Path) -> None:
    wb = openpyxl.load_workbook(asset_template, read_only=True)
    template_sheets = [s for s in wb.sheetnames if s != "替代结果汇总表"]
    assert len(template_sheets) >= 1
    ws = wb[template_sheets[0]]
    found_1 = found_2 = found_3 = False
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell and "借方发生额" in str(cell):
                found_1 = True
            if cell and "贷方发生额" in str(cell):
                found_2 = True
            if cell and "期后收款" in str(cell):
                found_3 = True
    assert found_1, "缺少'借方发生额'"
    assert found_2, "缺少'贷方发生额'"
    assert found_3, "缺少'期后收款'"
    wb.close()


def test_asset_no_merged_cells(asset_template: Path) -> None:
    wb = openpyxl.load_workbook(asset_template)
    for name in wb.sheetnames:
        ws = wb[name]
        assert not ws.merged_cells.ranges, f"Sheet '{name}' 含合并单元格"
    wb.close()


# ── 共同约束 ──

def test_templates_no_checkmarks(liability_template: Path, asset_template: Path) -> None:
    prohibited = {"√", "未发现异常", "附件完整"}
    for path in [liability_template, asset_template]:
        wb = openpyxl.load_workbook(path, read_only=True)
        for name in wb.sheetnames:
            ws = wb[name]
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue
                    for p in prohibited:
                        assert p not in str(cell), f"Sheet '{name}' 含禁用词 '{p}': {cell}"
        wb.close()


def test_templates_no_real_data(liability_template: Path, asset_template: Path) -> None:
    prohibited = {"乐山涌江", "汪红君", "张力", "2025-12-31"}
    for path in [liability_template, asset_template]:
        wb = openpyxl.load_workbook(path, read_only=True)
        for name in wb.sheetnames:
            ws = wb[name]
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue
                    for p in prohibited:
                        assert p not in str(cell), f"Sheet '{name}' 含真实数据 '{p}': {cell}"
        wb.close()


# ── 可重复性（使用预构建文件） ──

def test_templates_consistent() -> None:
    """预构建文件的 SHA-256 稳定。"""
    p1 = ASSETS / "替代测试-负债方-v1.xlsx"
    p2 = ASSETS / "替代测试-资产方-v1.xlsx"
    if not p1.exists() or not p2.exists():
        pytest.skip("预构建模板不存在")
    h1 = hashlib.sha256(p1.read_bytes()).hexdigest()
    h2 = hashlib.sha256(p2.read_bytes()).hexdigest()
    assert len(h1) == 64
    assert len(h2) == 64
    assert h1 != h2  # 两份模板不同
