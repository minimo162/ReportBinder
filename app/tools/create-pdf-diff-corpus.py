"""Generate deterministic, practical PDF pairs for visual-diff regression."""

from __future__ import annotations

import argparse
import json
import random
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas


def normalized_box(page, x, y, width, height):
    page_width, page_height = page
    return {
        "x": round(x / page_width, 6),
        "y": round((page_height - y - height) / page_height, 6),
        "width": round(width / page_width, 6),
        "height": round(height / page_height, 6),
    }


def page_frame(pdf, page, title):
    width, height = page
    pdf.setStrokeColor(colors.HexColor("#334155"))
    pdf.setFillColor(colors.HexColor("#0f172a"))
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(42, height - 48, title)
    pdf.setFont("Helvetica", 8)
    pdf.setFillColor(colors.HexColor("#64748b"))
    pdf.drawRightString(width - 42, 24, "ReportBinder visual-diff corpus")
    pdf.line(42, height - 58, width - 42, height - 58)


def text_only(path, changed):
    page = A4
    pdf = canvas.Canvas(str(path), pagesize=page, pageCompression=1)
    page_frame(pdf, page, "Department Activity Report")
    width, height = page
    lines = [
        "Purpose: consolidate monthly submissions into one review-ready document.",
        "Scope: Finance, Operations, Human Resources, and Compliance.",
        "All owners confirmed their source files before the reporting deadline.",
        "The first review found no missing mandatory attachments.",
        "The operating margin forecast remains within the approved range.",
        "The final release is scheduled for August 28." if changed else "The final release is scheduled for August 21.",
        "Reviewers should record material comments in the comparison history.",
        "Minor typography changes do not require a second approval cycle.",
        "Distribution is limited to the project members listed in the appendix.",
    ]
    pdf.setFillColor(colors.HexColor("#1f2937"))
    pdf.setFont("Times-Roman", 11)
    start_y = height - 92
    for index, line in enumerate(lines):
        pdf.drawString(54, start_y - index * 28, line)
    pdf.save()
    return [normalized_box(page, 48, start_y - 5 * 28 - 5, 360, 18)]


def resume_form(path, changed):
    page = A4
    pdf = canvas.Canvas(str(path), pagesize=page, pageCompression=1)
    page_frame(pdf, page, "Professional Profile")
    width, height = page
    left, right, top = 42, width - 42, height - 86
    row_heights = [42, 42, 78, 42, 115, 115]
    y = top
    labels = ["Name", "Contact", "Summary", "Availability", "Experience", "Qualifications"]
    values = [
        "Avery Morgan",
        "avery@example.test  /  +81 00 0000 0000",
        "Document operations specialist with cross-department reporting experience.",
        "Available from 2026-09-15" if changed else "Available from 2026-09-01",
        "2022-present  Reporting Operations\n2019-2022  Business Process Coordination",
        "Records management / PDF quality assurance / Office automation",
    ]
    pdf.setLineWidth(0.8)
    for label, value, row_height in zip(labels, values, row_heights):
        next_y = y - row_height
        pdf.rect(left, next_y, right - left, row_height)
        pdf.line(left + 112, next_y, left + 112, y)
        pdf.setFont("Helvetica-Bold", 9)
        pdf.drawString(left + 8, y - 18, label)
        pdf.setFont("Helvetica", 9)
        for line_index, line in enumerate(value.split("\n")):
            pdf.drawString(left + 122, y - 18 - line_index * 16, line)
        y = next_y
    pdf.save()
    availability_top = top - sum(row_heights[:3])
    return [normalized_box(page, left + 116, availability_top - row_heights[3] + 8, 210, 24)]


def complex_table(path, changed):
    page = A4
    pdf = canvas.Canvas(str(path), pagesize=page, pageCompression=1)
    page_frame(pdf, page, "Quarterly Control Matrix")
    width, height = page
    top, bottom = height - 105, height - 430
    columns = [42, 102, 260, 370, 455, width - 42]
    if changed:
        columns = [42, 102, 220, 300, 390, 470, width - 42]
    rows = [top - index * 32.5 for index in range(11)]
    pdf.setFillColor(colors.HexColor("#e2e8f0"))
    pdf.rect(columns[0], rows[1], columns[-1] - columns[0], rows[0] - rows[1], fill=1, stroke=0)
    pdf.setStrokeColor(colors.HexColor("#475569"))
    for x in columns:
        pdf.line(x, bottom, x, top)
    for y in rows:
        pdf.line(columns[0], y, columns[-1], y)
    headers = ["ID", "Control", "Owner", "Status", "Due"] if not changed else ["ID", "Control", "Evidence", "Owner", "Status", "Due"]
    pdf.setFillColor(colors.HexColor("#0f172a"))
    for index, header in enumerate(headers):
        pdf.setFont("Helvetica-Bold", 8)
        pdf.drawString(columns[index] + 5, rows[0] - 20, header)
    for row in range(1, 10):
        values = [f"C-{row:02d}", f"Review item {row}", f"Team {1 + row % 4}", "Open" if row % 3 else "Closed", f"Aug {10 + row}"]
        if changed:
            values = [values[0], values[1], f"E-{100 + row}", values[2], "Escalated" if row == 7 else values[3], values[4]]
        for column, value in enumerate(values):
            pdf.setFont("Helvetica", 7.5)
            pdf.drawString(columns[column] + 5, rows[row] - 20, value)
    pdf.setFont("Helvetica", 9)
    pdf.drawString(42, height - 470, "Approval note: all exceptions require evidence before publication.")
    if changed:
        pdf.drawString(42, height - 490, "Escalation owner: Internal Controls")
    pdf.save()
    return [
        normalized_box(page, 205, bottom - 3, 105, top - bottom + 6),
        normalized_box(page, 35, height - 500, 280, 22),
    ]


def visual_dashboard(path, changed):
    page = landscape(A4)
    pdf = canvas.Canvas(str(path), pagesize=page, pageCompression=1)
    page_frame(pdf, page, "Operations Dashboard")
    width, height = page
    colors_by_bar = ["#2563eb", "#0ea5e9", "#14b8a6", "#22c55e", "#84cc16"]
    heights = [92, 145, 118, 176 if changed else 132, 154]
    base_y = 125
    pdf.setStrokeColor(colors.HexColor("#94a3b8"))
    pdf.line(70, base_y, width - 60, base_y)
    for index, bar_height in enumerate(heights):
        x = 95 + index * 120
        pdf.setFillColor(colors.HexColor(colors_by_bar[index]))
        pdf.rect(x, base_y, 62, bar_height, fill=1, stroke=0)
        pdf.setFillColor(colors.HexColor("#334155"))
        pdf.setFont("Helvetica", 9)
        pdf.drawCentredString(x + 31, base_y - 18, f"Unit {index + 1}")
    pdf.setFillColor(colors.HexColor("#f1f5f9"))
    pdf.roundRect(width - 220, height - 180, 160, 70, 8, fill=1, stroke=0)
    pdf.setFillColor(colors.HexColor("#0f172a"))
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(width - 205, height - 140, "Completion 96%" if changed else "Completion 92%")
    pdf.save()
    return [
        normalized_box(page, 95 + 3 * 120 - 4, base_y - 4, 70, 188),
    ]


def page_alignment_document(path, logical_pages, design=1, number_footer=True):
    page = A4
    pdf = canvas.Canvas(str(path), pagesize=page, pageCompression=1)
    width, height = page
    for physical_page, logical_page in enumerate(logical_pages, start=1):
        if design == 1:
            page_frame(pdf, page, f"Controlled Procedure {logical_page}")
        else:
            pdf.setFillColor(colors.HexColor("#e0f2fe"))
            pdf.rect(0, height - 92, width, 92, fill=1, stroke=0)
            pdf.setFillColor(colors.HexColor("#0284c7"))
            pdf.rect(0, 0, 18, height, fill=1, stroke=0)
            pdf.setFillColor(colors.HexColor("#0c4a6e"))
            pdf.setFont("Helvetica-Bold", 16)
            pdf.drawString(42, height - 48, f"Enterprise Procedure {logical_page}")
            pdf.setFont("Helvetica", 8)
            pdf.drawString(42, height - 70, "Template 2026-B | Controlled copy")
        pdf.setFillColor(colors.HexColor("#0f172a"))
        pdf.setFont("Helvetica-Bold", 18)
        pdf.drawString(54, height - 145, f"Section {logical_page}")
        pdf.setFont("Helvetica", 10)
        for row in range(14):
            pdf.drawString(54, height - 185 - row * 28, f"{logical_page}-{row + 1:02d}  Evidence, owner, deadline and approval status for this control step.")
        # The logical page body remains identical when a page is inserted, but
        # the automatic physical-page footer changes. This is the practical case
        # that equal-index and whole-image-hash matching both get wrong.
        if number_footer:
            pdf.setFont("Helvetica", 8)
            pdf.drawCentredString(width / 2, 28, f"Page {physical_page}")
        pdf.showPage()
    pdf.save()


def scan_noise_document(path, seed, changed=False):
    page = A4
    width, height = page
    rng = random.Random(seed)
    pdf = canvas.Canvas(str(path), pagesize=page, pageCompression=1)
    pdf.setStrokeColor(colors.HexColor("#1e293b"))
    pdf.rect(36, 42, width - 72, height - 84, fill=0, stroke=1)
    pdf.setFillColor(colors.HexColor("#e2e8f0"))
    pdf.rect(36, height - 100, width - 72, 58, fill=1, stroke=1)
    pdf.setFillColor(colors.HexColor("#0f172a"))
    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(52, height - 77, "SITE INSPECTION RECORD")
    y = height - 155
    for row in range(10):
        pdf.rect(52, y - 29, width - 104, 34, fill=0, stroke=1)
        pdf.rect(62, y - 20, 14, 14, fill=0, stroke=1)
        pdf.setFont("Helvetica", 8)
        pdf.drawString(88, y - 16, f"Inspection item {row + 1:02d}: evidence, condition and corrective action")
        y -= 34
    if changed:
        pdf.setLineWidth(2.4)
        pdf.line(63, height - 169, 69, height - 177)
        pdf.line(69, height - 177, 79, height - 161)
        pdf.setLineWidth(1)
    # Isolated faint marks emulate scanner dust; each acquisition uses a
    # different seed and should remain below the connected-region threshold.
    pdf.setFillColor(colors.HexColor("#cbd5e1"))
    for _ in range(500):
        pdf.circle(rng.uniform(40, width - 40), rng.uniform(45, height - 105), 0.25, fill=1, stroke=0)
    pdf.save()


def build(output):
    output.mkdir(parents=True, exist_ok=True)
    cases = []
    for name, builder in [
        ("text-only", text_only),
        ("resume-form", resume_form),
        ("complex-table", complex_table),
        ("visual-dashboard", visual_dashboard),
    ]:
        before = output / f"{name}-before.pdf"
        after = output / f"{name}-after.pdf"
        builder(before, False)
        expected = builder(after, True)
        cases.append({"name": name, "before": before.name, "after": after.name, "page": 1, "expected": expected})
    alignment_before = output / "page-alignment-before.pdf"
    alignment_after = output / "page-alignment-after.pdf"
    page_alignment_document(alignment_before, ["A", "B", "C"])
    page_alignment_document(alignment_after, ["A", "INSERTED", "B", "C"])
    redesign_before = output / "page-redesign-before.pdf"
    redesign_after = output / "page-redesign-after.pdf"
    redesign_ambiguous = output / "page-redesign-ambiguous.pdf"
    page_alignment_document(redesign_before, ["A", "B", "C", "D"], design=1)
    page_alignment_document(redesign_after, ["A", "B", "INSERTED", "C", "D"], design=2)
    page_alignment_document(redesign_ambiguous, ["A", "INSERTED", "C", "D"], design=2)
    duplicate_before = output / "page-duplicate-before.pdf"
    duplicate_after = output / "page-duplicate-after.pdf"
    page_alignment_document(duplicate_before, ["A", "DUPLICATE", "DUPLICATE", "C"], number_footer=False)
    page_alignment_document(duplicate_after, ["A", "DUPLICATE", "DUPLICATE", "DUPLICATE", "C"], number_footer=False)
    scan_before = output / "scan-noise-before.pdf"
    scan_noise = output / "scan-noise-after.pdf"
    scan_changed = output / "scan-change-after.pdf"
    scan_noise_document(scan_before, 101, False)
    scan_noise_document(scan_noise, 202, False)
    scan_noise_document(scan_changed, 303, True)
    manifest = {
        "schemaVersion": 1,
        "dpi": 72,
        "minimumRecall": 1.0,
        "minimumPrecision": 0.6,
        "cases": cases,
        "pageAlignment": {
            "before": alignment_before.name,
            "after": alignment_after.name,
            "forward": [[1, 1], [0, 2], [2, 3], [3, 4]],
            "reverse": [[1, 1], [2, 0], [3, 2], [4, 3]],
        },
        "riskAlignment": {
            "redesignBefore": redesign_before.name,
            "redesignAfter": redesign_after.name,
            "redesignAmbiguous": redesign_ambiguous.name,
            "duplicateBefore": duplicate_before.name,
            "duplicateAfter": duplicate_after.name,
            "scanBefore": scan_before.name,
            "scanNoise": scan_noise.name,
            "scanChanged": scan_changed.name,
            "redesignForward": [[1, 1], [2, 2], [0, 3], [3, 4], [4, 5]],
            "redesignReverse": [[1, 1], [2, 2], [3, 0], [4, 3], [5, 4]],
        },
    }
    (output / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output / "manifest.json")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    build(parser.parse_args().output.resolve())
