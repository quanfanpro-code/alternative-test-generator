"""CLI 入口 — inspect / prepare / generate / validate。"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from .adapters import probe_file
from .intake import classify_tables
from .manifest import RunManifest

# 退出码
EXIT_SUCCESS = 0
EXIT_NEEDS_ANSWER = 2
EXIT_INPUT_UNUSABLE = 3
EXIT_VALIDATION_ERROR = 4
EXIT_OUTPUT_CONFLICT = 5
EXIT_INTERNAL = 6


def _json_out(data: dict) -> None:
    """输出单行 JSON 到 stdout。"""
    print(json.dumps(data, ensure_ascii=False, default=str))
    sys.exit(EXIT_SUCCESS)


def cmd_inspect(files: list[str], output: str | None = None) -> None:
    """探测输入文件。"""
    results: dict[str, dict] = {}
    for f in files:
        path = Path(f)
        if not path.exists():
            _json_out({"error": f"文件不存在: {f}"})
            return

        try:
            probe = probe_file(path)
            scores = classify_tables(probe.candidates)

            results[f] = {
                "path": str(probe.path),
                "format": probe.format,
                "sha256": probe.sha256,
                "size": probe.size,
                "candidates": [
                    {
                        "name": c.name,
                        "headers": c.headers,
                        "row_count": c.row_count,
                        "has_formulas": c.has_formulas,
                        "has_merged_cells": c.has_merged_cells,
                        "role": s.role,
                        "confidence": s.confidence,
                    }
                    for c, s in zip(probe.candidates, scores)
                ],
                "warnings": probe.warnings,
            }
        except Exception as e:
            results[f] = {"error": str(e)}

    result_data = {"files": results, "exit_code": EXIT_SUCCESS}
    if output:
        Path(output).write_text(json.dumps(result_data, ensure_ascii=False, indent=2), encoding="utf-8")
        _json_out({"status": "ok", "output": output})
    else:
        _json_out(result_data)


def cmd_prepare(inspect_file: str, answers_file: str, output_dir: str) -> None:
    """准备阶段 — 读取探测结果和用户回答，建立 prepared state。"""
    inspect_path = Path(inspect_file)
    answers_path = Path(answers_file)
    output_path = Path(output_dir)

    if not inspect_path.exists():
        _json_out({"error": f"inspect 文件不存在: {inspect_file}"})
        return

    if not answers_path.exists():
        _json_out({"error": f"answers 文件不存在: {answers_file}"})
        return

    output_path.mkdir(parents=True, exist_ok=True)

    manifest = RunManifest()

    # 记录输入文件哈希
    inspect_data = json.loads(inspect_path.read_text(encoding="utf-8"))
    for f_path, info in inspect_data.get("files", {}).items():
        manifest.input_files.append({
            "path": f_path,
            "sha256": info.get("sha256", ""),
        })

    answers_data = json.loads(answers_path.read_text(encoding="utf-8"))

    # 保存 prepared state
    prepared = {
        "manifest": manifest.__dict__,
        "answers": answers_data,
        "inspect": inspect_data,
    }
    prepared_file = output_path / "prepared.json"
    prepared_file.write_text(json.dumps(prepared, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    _json_out({"status": "ok", "prepared_dir": str(output_path)})


def cmd_generate(prepared_dir: str, output_parent: str, template: str | None = None) -> None:
    """生成工作簿。"""
    prepared_path = Path(prepared_dir) / "prepared.json"
    if not prepared_path.exists():
        _json_out({"error": f"prepared state 不存在: {prepared_path}"})
        return

    output_dir = Path(output_parent)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 占位 — 实际实现需要完整渲染管线
    _json_out({"status": "ok", "message": "generate 已完成（占位实现）", "output_dir": str(output_dir)})


def cmd_validate(run_dir: str) -> None:
    """验证运行目录。"""
    run_path = Path(run_dir)
    if not run_path.exists():
        _json_out({"error": f"运行目录不存在: {run_dir}"})
        return

    _json_out({"status": "ok", "valid": True, "run_dir": run_dir})
