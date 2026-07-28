"""模板生成器 — OfficeCLI 构建框架 + openpyxl 补尺寸，零合并单元格。"""

from __future__ import annotations

import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import yaml
from openpyxl.styles import Alignment, Border, Side


OC = "officecli"
THIN_BORDER_SIDE = Side(style="thin")


# ── 合同 ──

@dataclass
class TemplateSection:
    title: str
    normal_direction: str
    check_items: list[dict]


@dataclass
class TemplateDef:
    name: str
    summary_sheet: str
    template_sheet: str
    sections: list[TemplateSection]


@dataclass
class TemplateContract:
    version: str
    templates: dict[str, TemplateDef]
    required_sheet_names: list[str]
    prohibited_content: list[str]

    @classmethod
    def from_yaml(cls, path: Path) -> TemplateContract:
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f)
        templates = {}
        for key, td in data["templates"].items():
            sections = [TemplateSection(**s) for s in td["sections"]]
            templates[key] = TemplateDef(
                name=td["name"], summary_sheet=td["summary_sheet"],
                template_sheet=td["template_sheet"], sections=sections,
            )
        return cls(
            version=data["version"], templates=templates,
            required_sheet_names=data["required_sheet_names"],
            prohibited_content=data["prohibited_content"],
        )


# ── 列宽/行高配置 ──

_COL_WIDTHS = {"A": 3, "B": 6, "C": 12, "D": 8, "E": 10, "F": 28, "G": 14,
               "H": 14, "I": 3, "J": 12, "K": 10, "L": 12, "M": 10}
_MAIN_HEADERS = ["序号", "日    期", "凭证种类", "凭证号", "业务内容", "对方科目", "金    额"]


# ── 公开接口 ──

def build_liability_template(path: Path) -> Path:
    return _build_template(path, "liability")


def build_asset_template(path: Path) -> Path:
    return _build_template(path, "asset")


def build_default_template(output_dir: Path, side: str) -> Path:
    names = {"liability": "替代测试-负债方-v1.xlsx", "asset": "替代测试-资产方-v1.xlsx"}
    return _build_template(output_dir / names[side], side)


# ── 核心构建 ──

def _build_template(path: Path, side: str) -> Path:
    """OfficeCLI 构建模板框架 → openpyxl 补列宽行高。"""
    contract = TemplateContract.from_yaml(
        Path(__file__).resolve().parent.parent.parent
        / "assets" / "templates" / "template-contract.yaml"
    )
    tdef = contract.templates[side]
    path_str = str(path.resolve())

    # 清理旧文件
    _run(OC, "close", path_str)
    if path.exists():
        path.unlink()

    # Phase 1: OfficeCLI 构建
    _run(OC, "create", path_str, "--type", "workbook")
    _run(OC, "add", path_str, "/", "--type", "sheet", "--prop", f"name={tdef.summary_sheet}")
    _run(OC, "add", path_str, "/", "--type", "sheet", "--prop", f"name={tdef.template_sheet}")
    _run(OC, "remove", path_str, "/Sheet1")

    # 填汇总表
    _build_summary_sheet_oc(path_str, tdef, side)
    # 填明细模板
    _build_detail_sheet_oc(path_str, tdef, side)

    # 保存并关闭
    _run(OC, "save", path_str)
    _run(OC, "close", path_str)
    time.sleep(0.5)

    # Phase 2: openpyxl 补列宽行高 + 审计说明行高
    _apply_sizing(path)

    return path


def _run(*args: str) -> None:
    subprocess.run(list(args), capture_output=True, text=True, timeout=30)


# ── OfficeCLI 构建函数 ──

def _oc_set(file: str, cell: str, *props: str) -> None:
    """便捷：set 一个单元格。"""
    cmd = [OC, "set", file, cell]
    for p in props:
        cmd += ["--prop", p]
    _run(*cmd)


def _oc_cc(file: str, sheet: str, row: int, col_start: str, col_end: str) -> None:
    """设跨列居中：文字在最左格，其余格全设 centerContinuous。"""
    cols = _col_range(col_start, col_end)
    for col in cols:
        _oc_set(file, f"/{sheet}/{col}{row}", "alignment.horizontal=centerContinuous")


def _col_range(start: str, end: str) -> list[str]:
    """生成列字母序列，如 C 到 H → ['C','D','E','F','G','H']。"""
    result = []
    for c in range(ord(start), ord(end) + 1):
        result.append(chr(c))
    return result


def _build_summary_sheet_oc(file: str, tdef: TemplateDef, side: str) -> None:
    """OfficeCLI 构建汇总表。"""
    s = tdef.summary_sheet
    r = 2  # 从第2行开始

    # 标题：B2 放文字，C2-H2 设跨列居中
    _oc_set(file, f"/{s}/B{r}", "text=替代结果汇总表", "font.size=14", "font.bold=true")
    _oc_cc(file, s, r, "B", "H")
    r += 2

    # 信息行 R4
    for label, col in [("被审计单位：", "B"), ("", "C"), ("编制：", "E"), ("", "F"), ("日期：", "H"), ("", "I"), ("索引号：", "K"), ("", "L")]:
        if label:
            _oc_set(file, f"/{s}/{col}{r}", f"text={label}", "font.name=宋体", "font.size=10",
                    "border.top.style=thin", "border.bottom.style=thin",
                    "border.left.style=thin", "border.right.style=thin")
        else:
            _oc_set(file, f"/{s}/{col}{r}", "border.top.style=thin", "border.bottom.style=thin",
                    "border.left.style=thin", "border.right.style=thin")
    r += 1

    # 信息行 R5
    for label, col in [("报表截止日：", "B"), ("", "C"), ("复核：", "E"), ("", "F"), ("日期：", "H"), ("", "I"), ("项目：", "K"), ("", "L")]:
        if label:
            _oc_set(file, f"/{s}/{col}{r}", f"text={label}", "font.name=宋体", "font.size=10",
                    "border.top.style=thin", "border.bottom.style=thin",
                    "border.left.style=thin", "border.right.style=thin")
        else:
            _oc_set(file, f"/{s}/{col}{r}", "border.top.style=thin", "border.bottom.style=thin",
                    "border.left.style=thin", "border.right.style=thin")
    r += 2

    # 表头 R7: 索引号│债务人/债权人名称│币种│原币金额│通过替代测试可确认金额│备    注
    headers = ["索引号", "债务人/债权人名称", "币种", "原币金额", "通过替代测试可确认金额", "备    注"]
    for i, h in enumerate(headers):
        col = chr(ord("B") + i)
        _oc_set(file, f"/{s}/{col}{r}", f"text={h}", "font.bold=true", "font.size=10",
                "fill=C0C0C0", "border.all=thin")
        _oc_set(file, f"/{s}/{col}{r}", "alignment.horizontal=centerContinuous")
    r += 1

    # 合计行
    _oc_set(file, f"/{s}/B{r}", "text=共 0 行", "font.bold=true", "font.size=10", "border.all=thin")
    _oc_set(file, f"/{s}/C{r}", "border.all=thin")
    _oc_set(file, f"/{s}/D{r}", "border.all=thin")
    _oc_set(file, f"/{s}/E{r}", "border.all=thin")
    _oc_set(file, f"/{s}/F{r}", "border.all=thin")
    _oc_set(file, f"/{s}/G{r}", "border.all=thin")
    r += 2

    # 审计说明
    _oc_set(file, f"/{s}/B{r}", "text=审计说明：", "font.bold=true", "font.size=10")
    r += 1
    for col in _col_range("B", "G"):
        _oc_set(file, f"/{s}/{col}{r}", "border.all=thin")


def _build_detail_sheet_oc(file: str, tdef: TemplateDef, side: str) -> None:
    """OfficeCLI 构建明细模板。"""
    s = tdef.template_sheet
    title_text = "应付账款替代测试表" if side == "liability" else "应收账款替代测试表"
    r = 2

    # 标题 R2
    _oc_set(file, f"/{s}/B{r}", f"text={title_text}", "font.size=14", "font.bold=true")
    _oc_cc(file, s, r, "B", "L")
    r += 2

    # 信息行 R4
    for label, col in [("被审计单位：", "B"), ("", "C"), ("编制：", "E"), ("", "F"), ("日期：", "H"), ("", "I"), ("索引号：", "K"), ("", "L")]:
        props = [f"text={label}", "font.name=宋体", "font.size=10"] if label else ["font.size=10"]
        props += ["border.top.style=thin", "border.bottom.style=thin",
                  "border.left.style=thin", "border.right.style=thin"]
        _oc_set(file, f"/{s}/{col}{r}", *props)
    r += 1

    # 信息行 R5
    for label, col in [("报表截止日：", "B"), ("", "C"), ("复核：", "E"), ("", "F"), ("日期：", "H"), ("", "I"), ("项目：", "K"), ("", "L")]:
        props = [f"text={label}", "font.name=宋体", "font.size=10"] if label else ["font.size=10"]
        props += ["border.top.style=thin", "border.bottom.style=thin",
                  "border.left.style=thin", "border.right.style=thin"]
        _oc_set(file, f"/{s}/{col}{r}", *props)
    r += 2

    # 选择账户行
    account_label = "供应商" if side == "liability" else "客户"
    _oc_set(file, f"/{s}/B{r}", f"text=选择需要测试的账户：", "font.size=10")
    r += 2

    # 三个板块
    for section in tdef.sections:
        r = _build_section_oc(file, s, r, section)
        r += 1

    # 审计说明
    r += 1
    _oc_set(file, f"/{s}/B{r}", "text=审计说明：", "font.bold=true", "font.size=10")
    r += 1
    end_col = "L" if side == "liability" else "K"
    for col in _col_range("B", end_col):
        _oc_set(file, f"/{s}/{col}{r}", "border.all=thin")


def _build_section_oc(file: str, sheet: str, start_row: int, section: TemplateSection) -> int:
    """OfficeCLI 构建一个版块，返回下一可用行号。"""
    r = start_row

    # 板块标题
    _oc_set(file, f"/{sheet}/B{r}", f"text={section.title}", "font.bold=true", "font.size=11",
            "fill=C0C0C0")
    check_cols = sum(len(item["sub_columns"]) for item in section.check_items)
    last_data_col = 7 + check_cols  # B-H(7列记账) + 检查列
    last_col_letter = chr(ord("B") + last_data_col - 1)
    _oc_cc(file, sheet, r, "B", last_col_letter)
    r += 1

    # 表头第一层：记账凭证 + 各检查材料名
    _oc_set(file, f"/{sheet}/B{r}", "text=记账凭证", "font.bold=true", "font.size=10")
    _oc_cc(file, sheet, r, "B", "H")
    col_offset = 9  # I 列开始
    for item in section.check_items:
        start_c = chr(ord("A") + col_offset)
        end_c = chr(ord("A") + col_offset + len(item["sub_columns"]) - 1)
        _oc_set(file, f"/{sheet}/{start_c}{r}", f"text={item['name']}", "font.bold=true", "font.size=10")
        _oc_cc(file, sheet, r, start_c, end_c)
        col_offset += len(item["sub_columns"])
    r += 1

    # 表头第二层
    for i, h in enumerate(_MAIN_HEADERS):
        col = chr(ord("B") + i)
        _oc_set(file, f"/{sheet}/{col}{r}", f"text={h}", "font.bold=true", "font.size=10",
                "fill=C0C0C0", "border.all=thin")
    col_offset = 9
    for item in section.check_items:
        for sub in item["sub_columns"]:
            col = chr(ord("A") + col_offset)
            _oc_set(file, f"/{sheet}/{col}{r}", f"text={sub}", "font.bold=true", "font.size=10",
                    "fill=C0C0C0", "border.all=thin")
            col_offset += 1
    r += 1

    # 小计行
    _oc_set(file, f"/{sheet}/B{r}", "text=0", "font.bold=true", "font.size=10", "border.all=thin")
    return r


# ── openpyxl 补尺寸 ──

def _apply_sizing(path: Path) -> None:
    """openpyxl 补充列宽和行高。"""
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        for col_letter, width in _COL_WIDTHS.items():
            ws.column_dimensions[col_letter].width = width
        # 审计说明行高
        if ws.max_row > 10:
            last_row = ws.max_row
            if ws.cell(row=last_row - 1, column=2).value == "审计说明：":
                ws.row_dimensions[last_row].height = 60
            # 第二个审计说明行（如果存在）
            elif ws.max_row > 5:
                for r in range(ws.max_row - 5, ws.max_row + 1):
                    val = ws.cell(row=r, column=2).value
                    if val and "审计说明" in str(val):
                        if r + 1 <= ws.max_row:
                            ws.row_dimensions[r + 1].height = 60
                        break
    wb.save(path)
    wb.close()


# ── 验证 ──

@dataclass
class TemplateValidation:
    valid: bool
    issues: list[str] = field(default_factory=list)


def validate_template(path: Path, contract: TemplateContract) -> TemplateValidation:
    issues = []
    wb = openpyxl.load_workbook(path, read_only=True)
    sheet_names = wb.sheetnames

    for req in contract.required_sheet_names:
        if req not in sheet_names:
            issues.append(f"缺少 Sheet: {req}")

    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                cell_str = str(cell)
                for disallowed in contract.prohibited_content:
                    if disallowed in cell_str:
                        issues.append(f"Sheet '{sheet_name}' 含禁用内容 '{disallowed}'")
                        break

    wb.close()

    wb = openpyxl.load_workbook(path)
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        if ws.merged_cells.ranges:
            issues.append(f"Sheet '{sheet_name}' 含合并单元格: {ws.merged_cells.ranges}")
    wb.close()

    return TemplateValidation(valid=len(issues) == 0, issues=issues)
