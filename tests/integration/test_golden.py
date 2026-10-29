"""以实际成品回归验证取代旧格式的静态占位样本。"""
import openpyxl
from alternative_test.fileio import calc_sha256
from alternative_test.renderer import render_workbook, validate_workbook
from tests.unit.test_生成验收 import context, TEMPLATE

def test_template_stays_unchanged_and_output_is_valid(tmp_path):
    before = calc_sha256(TEMPLATE)
    output = render_workbook(context(8), TEMPLATE, tmp_path / "回归样本.xlsx")
    assert calc_sha256(TEMPLATE) == before
    assert validate_workbook(output)["valid"]
    w = openpyxl.load_workbook(output)
    s = w["演示供应商有限公司"]
    assert s["H20"].value == "=SUM(H12:H19)"
    assert s["B23"].value == "二、借方发生额"
    w.close()
