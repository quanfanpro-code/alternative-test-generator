"""包基础测试：导入、版本、隔离扫描。"""

from __future__ import annotations

from pathlib import Path

import alternative_test


class TestPackageBasics:
    """测试包可正常导入且版本正确。"""

    def test_version_exists(self) -> None:
        assert alternative_test.__version__
        assert isinstance(alternative_test.__version__, str)

    def test_version_format(self) -> None:
        parts = alternative_test.__version__.split(".")
        assert len(parts) == 3
        for p in parts:
            assert p.isdigit()


class TestLegacyIsolation:
    """验证正式目录不含旧资产。"""

    def test_no_xls_outside_fixtures(self) -> None:
        """除 tests/fixtures/ 外不得存在 .xls 文件。"""
        root = Path(__file__).resolve().parents[2]
        for f in root.rglob("*.xls"):
            if f.is_relative_to(root / "tests" / "fixtures"):
                continue
            if f.is_relative_to(root / ".venv"):
                continue
            assert False, f"发现不应存在的 .xls 文件: {f}"
