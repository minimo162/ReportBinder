#!/usr/bin/env python3
"""Exercise PR 7 structured PDF composition with real multi-page PDFs."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from pathlib import Path

from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[2]
PDFBOX = ROOT / "app" / "lib" / "pdfbox"


def make_source(path: Path, label: str, pages: int) -> None:
    doc = canvas.Canvas(str(path), pagesize=A4)
    for number in range(1, pages + 1):
        doc.setFont("Helvetica-Bold", 24)
        doc.drawString(72, 760, f"{label} page {number}")
        doc.setFont("Helvetica", 11)
        doc.drawString(72, 730, f"source marker: {label.lower()}-{number}")
        doc.showPage()
    doc.save()


def outline_titles(reader: PdfReader) -> list[str]:
    result: list[str] = []

    def walk(nodes: list[object]) -> None:
        for node in nodes:
            if isinstance(node, list):
                walk(node)
            else:
                title = getattr(node, "title", "")
                if title:
                    result.append(str(title))

    walk(reader.outline)
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--keep", action="store_true")
    args = parser.parse_args()
    work = ROOT / "tmp" / "pdfs" / "pr7-final-composition"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    source_a = work / "alpha.pdf"
    source_b = work / "beta.pdf"
    output = work / "structured.pdf"
    manifest_path = work / "manifest.json"
    make_source(source_a, "Alpha", 3)
    make_source(source_b, "Beta", 2)
    manifest = {
        "schemaVersion": 3,
        "projectId": "PR7",
        "createdAt": "2026-08-06T23:00:00+09:00",
        "outputPdf": str(output),
        "document": {
            "title": "実践資料パック",
            "subtitle": "PR 7 structured composition",
            "targetName": "本体",
            "includeCover": True,
            "includeToc": True,
            "includeSectionDividers": True,
        },
        "pageNumber": {"fontSize": 8, "bottomPt": 18, "format": "hyphenated"},
        "pages": [
            {
                "pageId": "alpha-item",
                "title": "Alpha excerpt",
                "bookmarkTitle": "営業部 / Alpha excerpt",
                "sectionId": "source:alpha",
                "sectionTitle": "営業部",
                "sourcePdf": str(source_a),
                "sourcePageStart": 2,
                "sourcePageEnd": 3,
                "numberingMode": "visible",
                "punchShiftPt": 0,
            },
            {
                "pageId": "beta-item",
                "title": "Beta first page",
                "bookmarkTitle": "管理部 / Beta first page",
                "sectionId": "source:beta",
                "sectionTitle": "管理部",
                "sourcePdf": str(source_b),
                "sourcePageStart": 1,
                "sourcePageEnd": 1,
                "numberingMode": "visible",
                "punchShiftPt": 0,
            },
        ],
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    java_candidates = [ROOT / "app" / "lib" / "java" / "bin" / "java.exe"]
    java_candidates.extend((ROOT / "app" / "lib" / "java").glob("**/bin/java.exe"))
    java = next((candidate for candidate in java_candidates if candidate.is_file()), None)
    if java is None:
        java_name = shutil.which("java")
        if not java_name:
            raise RuntimeError("Java runtime not found")
        java = Path(java_name)
    command = [
        str(java),
        "-cp",
        f"{PDFBOX / 'ReportPdfComposer.jar'};{PDFBOX / 'pdfbox-app.jar'}",
        "ReportPdfComposer",
        "--manifest",
        str(manifest_path),
    ]
    result = subprocess.run(command, cwd=ROOT, text=True, capture_output=True, encoding="utf-8", errors="replace")
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    reader = PdfReader(output)
    assert len(reader.pages) == 7, f"expected 7 physical pages, got {len(reader.pages)}"
    assert "Alpha page 2" in (reader.pages[3].extract_text() or "")
    assert "Alpha page 3" in (reader.pages[4].extract_text() or "")
    assert "Alpha page 1" not in "\n".join((p.extract_text() or "") for p in reader.pages)
    assert "Beta page 1" in (reader.pages[6].extract_text() or "")
    assert outline_titles(reader) == ["営業部 / Alpha excerpt", "管理部 / Beta first page"]
    print(f"PASS structured composition: {output}")
    print("PASS pageRange selected Alpha 2-3 and Beta 1")
    print("PASS cover, TOC, section dividers, physical page count=7")
    print("PASS item-derived PDF bookmarks")
    if not args.keep:
        shutil.rmtree(work)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
