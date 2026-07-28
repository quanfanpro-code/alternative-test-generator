#! python3.14
"""构建默认脱敏模板 — 负债方 + 资产方各一份。"""

from pathlib import Path

from alternative_test.templates import build_liability_template, build_asset_template

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets" / "templates"

if __name__ == "__main__":
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    p1 = build_liability_template(ASSETS_DIR / "替代测试-负债方-v1.xlsx")
    p2 = build_asset_template(ASSETS_DIR / "替代测试-资产方-v1.xlsx")
    print(f"✅ 负债方模板: {p1}")
    print(f"✅ 资产方模板: {p2}")
