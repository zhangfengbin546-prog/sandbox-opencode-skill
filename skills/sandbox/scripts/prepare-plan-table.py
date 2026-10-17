# -*- coding: utf-8 -*-
"""Prepare spreadsheet deliverables without changing workbook contents."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
from pathlib import Path


SUPPORTED = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".csv"}


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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare sandbox plan table deliverables")
    parser.add_argument("--input", action="append", required=True, help="Repeat for every plan table")
    parser.add_argument(
        "--output-dir",
        help="Output directory; defaults to the current user's desktop",
    )
    return parser.parse_args()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def stable_output(output_dir: Path, source: Path, used_names: set[str]) -> Path:
    candidate = output_dir / source.name
    index = 2
    while candidate.exists() or candidate.name.casefold() in used_names:
        candidate = output_dir / f"{source.stem}_{index}{source.suffix}"
        index += 1
    used_names.add(candidate.name.casefold())
    return candidate


def main() -> None:
    args = parse_args()
    output_dir = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else desktop_directory()
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    outputs: list[dict[str, str]] = []
    used_names: set[str] = set()
    for raw in args.input:
        source = Path(raw).expanduser().resolve()
        if not source.is_file():
            raise FileNotFoundError(f"方案表不存在：{source}")
        if source.suffix.lower() not in SUPPORTED:
            raise ValueError(f"不支持的表格格式：{source.suffix}")

        output = stable_output(output_dir, source, used_names)
        source_hash = sha256(source)
        source_size = source.stat().st_size
        if source != output:
            shutil.copy2(source, output)
        if source_hash != sha256(output) or source_size != output.stat().st_size:
            output.unlink(missing_ok=True)
            raise IOError(f"方案表准备失败：{source.name}")
        if not output.is_file():
            raise IOError(f"方案表未生成：{output}")
        outputs.append(
            {
                "name": output.name,
                "path": str(output),
                "uri": output.as_uri(),
            }
        )

    print(json.dumps({"outputs": outputs}, ensure_ascii=False))


if __name__ == "__main__":
    main()
