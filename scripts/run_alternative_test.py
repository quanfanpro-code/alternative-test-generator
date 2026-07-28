#! python3.14
"""CLI 入口脚本 — alternative-test-generator。"""

import sys
from pathlib import Path

# 将 src 加入 path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from alternative_test.cli import (
    cmd_generate,
    cmd_inspect,
    cmd_prepare,
    cmd_validate,
)


def main() -> None:
    if len(sys.argv) < 2:
        print("用法: run_alternative_test.py <inspect|prepare|generate|validate> [选项]")
        sys.exit(6)

    cmd = sys.argv[1]

    if cmd == "inspect":
        # python run_alternative_test.py inspect file1.xlsx file2.xlsx --output inspect.json
        args = sys.argv[2:]
        output = None
        if "--output" in args:
            idx = args.index("--output")
            output = args[idx + 1]
            args = args[:idx] + args[idx + 2:]
        cmd_inspect(args, output)

    elif cmd == "prepare":
        # python run_alternative_test.py prepare --inspect inspect.json --answers answers.json --output prepared/
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--inspect", required=True)
        parser.add_argument("--answers", required=True)
        parser.add_argument("--output", required=True)
        ns = parser.parse_args(sys.argv[2:])
        cmd_prepare(ns.inspect, ns.answers, ns.output)

    elif cmd == "generate":
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--prepared", required=True)
        parser.add_argument("--output-parent", required=True)
        parser.add_argument("--template", default=None)
        ns = parser.parse_args(sys.argv[2:])
        cmd_generate(ns.prepared, ns.output_parent, ns.template)

    elif cmd == "validate":
        if len(sys.argv) < 3:
            print("用法: run_alternative_test.py validate <run-dir>")
            sys.exit(6)
        cmd_validate(sys.argv[2])

    else:
        print(f"未知命令: {cmd}")
        sys.exit(6)


if __name__ == "__main__":
    main()
