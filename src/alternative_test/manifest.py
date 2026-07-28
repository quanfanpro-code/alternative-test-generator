"""运行清单 — 记录输入输出、哈希、用户决定和验证结果。"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class RunManifest:
    """运行清单。"""
    run_id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])
    version: str = "2.0.0"
    input_files: list[dict[str, str]] = field(default_factory=list)
    output_file: str = ""
    decision_log: list[dict[str, Any]] = field(default_factory=list)
    controls_summary: list[dict[str, Any]] = field(default_factory=list)
    downgrades: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    user_confirmations: list[str] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(self.__dict__, ensure_ascii=False, indent=2, default=str)

    def save(self, path: Path) -> None:
        path.write_text(self.to_json(), encoding="utf-8")
