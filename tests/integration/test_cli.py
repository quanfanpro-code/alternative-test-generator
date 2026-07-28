"""Task 11 — CLI 集成测试。"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

RUNNER = str(Path(__file__).resolve().parents[2] / "scripts" / "run_alternative_test.py")


def _run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, RUNNER, *args],
        capture_output=True,
        text=True,
        timeout=30,
    )


# ── inspect ──

def test_inspect_xlsx(tmp_path: Path) -> None:
    """inspect 命令探测 xlsx 文件。"""
    from tests.factories import make_simple_balance_xlsx
    xlsx_path = tmp_path / "test.xlsx"
    make_simple_balance_xlsx(xlsx_path)

    result = _run("inspect", str(xlsx_path))
    assert result.returncode == 0

    data = json.loads(result.stdout.strip())
    assert "files" in data
    assert str(xlsx_path) in data["files"]
    file_info = data["files"][str(xlsx_path)]
    assert file_info["format"] == "xlsx"


def test_inspect_output_file(tmp_path: Path) -> None:
    """inspect --output 写入文件。"""
    from tests.factories import make_simple_balance_xlsx
    xlsx_path = tmp_path / "test.xlsx"
    make_simple_balance_xlsx(xlsx_path)

    output = tmp_path / "inspect.json"
    result = _run("inspect", str(xlsx_path), "--output", str(output))
    assert result.returncode == 0
    assert output.exists()


# ── prepare ──

def test_prepare_creates_state(tmp_path: Path) -> None:
    """prepare 创建 prepared state。"""
    from tests.factories import make_simple_balance_xlsx
    xlsx_path = tmp_path / "test.xlsx"
    make_simple_balance_xlsx(xlsx_path)

    # 先 inspect
    inspect_out = tmp_path / "inspect.json"
    _run("inspect", str(xlsx_path), "--output", str(inspect_out))

    # 创建 answers
    answers = tmp_path / "answers.json"
    answers.write_text(json.dumps({
        "period_start": "2025-01-01",
        "period_end": "2025-12-31",
        "entity_ids": ["E001"],
    }))

    # prepare
    prep_dir = tmp_path / "prepared"
    result = _run("prepare", "--inspect", str(inspect_out), "--answers", str(answers), "--output", str(prep_dir))
    assert result.returncode == 0

    prep_file = prep_dir / "prepared.json"
    assert prep_file.exists()


# ── generate（占位） ──

def test_generate_placeholder(tmp_path: Path) -> None:
    """generate 命令（占位实现）正常退出。"""
    prep_dir = tmp_path / "prepared"
    prep_dir.mkdir()
    (prep_dir / "prepared.json").write_text("{}")

    result = _run("generate", "--prepared", str(prep_dir), "--output-parent", str(tmp_path))
    assert result.returncode == 0


# ── validate ──

def test_validate(tmp_path: Path) -> None:
    """validate 命令检查目录。"""
    result = _run("validate", str(tmp_path))
    assert result.returncode == 0
