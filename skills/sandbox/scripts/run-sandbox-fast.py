# -*- coding: utf-8 -*-
"""Run extraction, deterministic strategy generation, and workbook delivery once."""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path


TABLE_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".csv"}


def desktop_directory() -> Path:
    """Resolve the current user's configured desktop, including redirected desktops."""
    if os.name == "nt":
        try:
            import winreg

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                raw, _ = winreg.QueryValueEx(key, "Desktop")
            return Path(os.path.expandvars(raw)).expanduser().resolve()
        except (OSError, ImportError):
            pass
    return (Path.home() / "Desktop").resolve()


def force_utf8_console() -> None:
    """Keep OpenCode tool output valid on Windows hosts using code page 936."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the fast sandbox workflow")
    parser.add_argument("--rule", action="append", required=True, help="Repeat for rule files")
    parser.add_argument("--plan", action="append", help="Repeat for user-supplied plan tables")
    parser.add_argument("--role", required=True, choices=("CEO", "财务", "生产", "人力"))
    parser.add_argument("--style", required=True, choices=("保守型", "激进型"))
    parser.add_argument("--makeup-market", required=True, choices=("是", "否"))
    parser.add_argument(
        "--emit-strategy",
        action="store_true",
        help="Print the generated Markdown between stable markers",
    )
    parser.add_argument(
        "--output-dir",
        help="Workbook delivery directory; defaults to the current user's desktop",
    )
    return parser.parse_args()


def run(command: list[str]) -> str:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(
        command,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
    )
    return completed.stdout.strip()


def main() -> None:
    force_utf8_console()
    started = time.perf_counter()
    args = parse_args()
    scripts = Path(__file__).resolve().parent
    skill_root = scripts.parent
    delivery_dir = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else desktop_directory()
    )
    delivery_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="sandbox-strategy-") as temporary:
        work_dir = Path(temporary).resolve()
        rules_source = work_dir / "rules-source.txt"
        extract_report = work_dir / "extract-report.json"
        strategy_path = work_dir / "企业沙盘战略方案.md"

        extract_command = [
            sys.executable,
            str(scripts / "extract-rules-fast.py"),
        ]
        for rule in args.rule:
            extract_command.extend(["--input", str(Path(rule).expanduser().resolve())])
        extract_command.extend(["--output", str(rules_source), "--report", str(extract_report)])
        run(extract_command)

        run(
            [
                sys.executable,
                str(scripts / "generate-strategy-fast.py"),
                "--rules-source",
                str(rules_source),
                "--role",
                args.role,
                "--style",
                args.style,
                "--makeup-market",
                args.makeup_market,
                "--output",
                str(strategy_path),
            ]
        )
        if not strategy_path.is_file():
            raise IOError(f"战略方案未生成：{strategy_path}")
        strategy_text = strategy_path.read_text(encoding="utf-8")

    plans = [Path(item).expanduser().resolve() for item in (args.plan or [])]
    if not plans:
        plans = sorted(
            path.resolve()
            for path in (skill_root / "assets").iterdir()
            if path.is_file() and path.suffix.lower() in TABLE_SUFFIXES
        )
    if not plans:
        raise FileNotFoundError("没有找到方案表；请上传表格或在 assets 中放入表格")

    prepare_command = [sys.executable, str(scripts / "prepare-plan-table.py")]
    for plan in plans:
        prepare_command.extend(["--input", str(plan)])
    prepare_command.extend(["--output-dir", str(delivery_dir)])
    prepared = json.loads(run(prepare_command))

    for item in prepared["outputs"]:
        if not Path(item["path"]).is_file():
            raise IOError(f"方案表未生成：{item['path']}")

    result = {
        "status": "ok",
        "strategy": {
            "characters": len(strategy_text),
        },
        "workbooks": prepared["outputs"],
        "delivery_dir": str(delivery_dir),
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }
    if args.emit_strategy:
        print("---SANDBOX-STRATEGY-BEGIN---")
        print(strategy_text, end="")
        print("---SANDBOX-STRATEGY-END---")
    # ASCII JSON remains readable even when a Windows host uses a legacy console code page.
    print(json.dumps(result, ensure_ascii=True))


if __name__ == "__main__":
    main()
