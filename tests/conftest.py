"""测试配置 — 共享 fixtures 和 --refresh-golden 选项。"""

from pathlib import Path

import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--refresh-golden",
        action="store_true",
        default=False,
        help="重新生成 golden 文件",
    )


@pytest.fixture
def data_dir() -> Path:
    """测试数据目录。"""
    return Path(__file__).resolve().parent


@pytest.fixture
def fixtures_dir(data_dir: Path) -> Path:
    """生成 fixture 目录。"""
    d = data_dir / "fixtures" / "generated"
    d.mkdir(parents=True, exist_ok=True)
    return d


@pytest.fixture
def golden_dir(data_dir: Path) -> Path:
    """golden 文件目录。"""
    d = data_dir / "golden"
    d.mkdir(parents=True, exist_ok=True)
    return d
