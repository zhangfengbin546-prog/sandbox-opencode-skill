# -*- coding: utf-8 -*-
"""Fast local text extraction for sandbox rule files.

Supported: PDF with a text layer, DOCX, legacy DOC through installed Word,
TXT/MD/JSON/CSV, and XLSX. Produces one UTF-8 text file plus a JSON report.
It never executes document content.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract sandbox competition rules locally")
    parser.add_argument("--input", required=True, action="append", help="Repeat for multiple files")
    parser.add_argument("--output", required=True, help="Combined UTF-8 text")
    parser.add_argument("--report", help="JSON extraction report")
    return parser.parse_args()


def paragraph_text(node: ET.Element) -> str:
    return "".join(part.text or "" for part in node.iter(W + "t")).strip()


def table_text(table: ET.Element) -> list[str]:
    lines = ["[TABLE]"]
    for row in table.findall(W + "tr"):
        cells = []
        for cell in row.findall(W + "tc"):
            pieces = [paragraph_text(p) for p in cell.iter(W + "p")]
            cells.append(" ".join(piece for piece in pieces if piece))
        lines.append(" | ".join(cells))
    lines.append("[/TABLE]")
    return lines


def extract_docx(path: Path) -> str:
    lines: list[str] = []
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        body = root.find(W + "body")
        if body is not None:
            for child in body:
                if child.tag == W + "p":
                    text = paragraph_text(child)
                    if text:
                        lines.append(text)
                elif child.tag == W + "tbl":
                    lines.extend(table_text(child))
        for optional in ["word/footnotes.xml", "word/endnotes.xml"]:
            if optional in archive.namelist():
                lines.append(f"[{Path(optional).stem.upper()}]")
                part = ET.fromstring(archive.read(optional))
                for paragraph in part.iter(W + "p"):
                    text = paragraph_text(paragraph)
                    if text:
                        lines.append(text)
        for name in sorted(n for n in archive.namelist() if re.fullmatch(r"word/(header|footer)\d+\.xml", n)):
            part = ET.fromstring(archive.read(name))
            content = [paragraph_text(p) for p in part.iter(W + "p")]
            content = [text for text in content if text]
            if content:
                lines.append(f"[{Path(name).stem.upper()}]")
                lines.extend(content)
    return "\n".join(lines)


def extract_pdf(path: Path) -> tuple[str, list[str]]:
    import fitz

    warnings: list[str] = []
    document = fitz.open(path)
    pages = []
    for index, page in enumerate(document, start=1):
        text = page.get_text("text", sort=True).strip()
        pages.append(f"[PAGE {index}]\n{text}")
    combined = "\n\n".join(pages)
    nonspace = len(re.sub(r"\s+", "", combined))
    if len(document) and nonspace < max(100, len(document) * 30):
        warnings.append("PDF 文本层过少，可能是扫描件；三分钟快速路径无法保证完整 OCR")
    return combined, warnings


def extract_doc_with_word(path: Path) -> str:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    temp_path = Path(tempfile.gettempdir()) / f"sandbox-{os.getpid()}-{time.time_ns()}.txt"
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        document.SaveAs2(str(temp_path), FileFormat=7, Encoding=65001, AddToRecentFiles=False)
        document.Close(False)
        document = None
        return temp_path.read_text(encoding="utf-8-sig", errors="replace")
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
        temp_path.unlink(missing_ok=True)


def extract_xlsx(path: Path) -> str:
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=False, read_only=True)
    lines: list[str] = []
    for sheet_name in workbook.sheetnames:
        lines.append(f"[SHEET {sheet_name}]")
        for row in workbook[sheet_name].iter_rows(values_only=True):
            values = ["" if value is None else str(value) for value in row]
            if any(value.strip() for value in values):
                while values and not values[-1].strip():
                    values.pop()
                lines.append(" | ".join(values))
    return "\n".join(lines)


def extract_one(path: Path) -> tuple[str, list[str]]:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".docx":
        return extract_docx(path), []
    if suffix == ".doc":
        return extract_doc_with_word(path), []
    if suffix in {".xlsx", ".xlsm"}:
        return extract_xlsx(path), []
    if suffix in {".txt", ".md", ".json", ".csv"}:
        return path.read_text(encoding="utf-8-sig", errors="replace"), []
    raise ValueError(f"不支持的规则文件格式：{path.suffix}")


def main() -> None:
    started = time.perf_counter()
    args = parse_args()
    output = Path(args.output).resolve()
    report_path = Path(args.report).resolve() if args.report else output.with_suffix(output.suffix + ".json")
    sections, records, warnings = [], [], []
    for raw in args.input:
        path = Path(raw).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"规则文件不存在：{path}")
        text, item_warnings = extract_one(path)
        sections.append(f"[FILE {path.name}]\n{text.strip()}")
        records.append({"path": str(path), "characters": len(text), "warnings": item_warnings})
        warnings.extend(f"{path.name}: {warning}" for warning in item_warnings)
    combined = "\n\n".join(sections).strip() + "\n"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(combined, encoding="utf-8")
    report = {
        "status": "ok",
        "output": str(output),
        "files": records,
        "characters": len(combined),
        "warnings": warnings,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
