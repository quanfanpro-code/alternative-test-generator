"""正式入口必须正确报告成功和失败，不能把空输入视为成功。"""
import json
import os
import subprocess
import sys
from pathlib import Path

RUNNER = Path(__file__).resolve().parents[2] / "scripts/run_alternative_test.py"

def run(*args):
    env = dict(os.environ, PYTHONUTF8="1", PYTHONDONTWRITEBYTECODE="1")
    return subprocess.run([sys.executable, "-B", "-X", "utf8", str(RUNNER), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", timeout=30, env=env)

def test_inspect_csv(tmp_path):
    source = tmp_path / "测试.csv"
    source.write_text("供应商,期末余额\n演示单位,100\n", encoding="utf-8-sig")
    result = run("inspect", source)
    assert result.returncode == 0
    assert json.loads(result.stdout)["files"][str(source)]["format"] == "csv"

def test_missing_input_is_error(tmp_path):
    assert run("inspect", tmp_path / "不存在.xlsx").returncode == 3

def test_incomplete_prepare_is_error(tmp_path):
    info, answers = tmp_path / "探测.json", tmp_path / "答案.json"
    info.write_text('{"files":{}}', encoding="utf-8")
    answers.write_text('{}', encoding="utf-8")
    assert run("prepare", "--inspect", info, "--answers", answers, "--output", tmp_path / "准备").returncode == 3
    assert not (tmp_path / "准备/prepared.json").exists()

def test_incomplete_generate_is_error(tmp_path):
    prepared = tmp_path / "prepared.json"
    prepared.write_text('{}', encoding="utf-8")
    assert run("generate", "--prepared", tmp_path, "--output-parent", tmp_path / "输出").returncode == 3
    assert not list(tmp_path.glob("*.xlsx"))

def test_empty_run_dir_cannot_pass_validation(tmp_path):
    assert run("validate", tmp_path).returncode == 4
