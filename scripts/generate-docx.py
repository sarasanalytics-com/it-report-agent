#!/usr/bin/env python3
"""
Convert a Markdown report to a Word (.docx) document.

Usage:
    python generate-docx.py <markdown_file> <output_docx>

Example:
    python generate-docx.py output/full-report.md output/IT-Weekly-Report.docx
"""

import re
import sys
import pathlib

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


def add_table_from_md(doc, header_line: str, rows: list[str]) -> None:
    """Parse markdown table lines and add a Word table."""
    def parse_cells(line: str) -> list[str]:
        return [c.strip() for c in line.strip().strip("|").split("|")]

    headers = parse_cells(header_line)
    num_cols = len(headers)

    table = doc.add_table(rows=1, cols=num_cols, style="Light Grid Accent 1")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    # Data rows
    for row_line in rows:
        cells = parse_cells(row_line)
        row = table.add_row()
        for i, val in enumerate(cells[:num_cols]):
            row.cells[i].text = val
            for p in row.cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)


def md_to_docx(md_text: str, output_path: pathlib.Path) -> None:
    """Convert markdown text to a Word document."""
    doc = Document()

    # Set default font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(10)

    lines = md_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Skip empty lines
        if not stripped:
            i += 1
            continue

        # Horizontal rule
        if stripped in ("---", "***", "___"):
            doc.add_paragraph("").runs  # spacer
            i += 1
            continue

        # Headings
        if stripped.startswith("# "):
            p = doc.add_heading(stripped[2:], level=1)
            i += 1
            continue
        if stripped.startswith("## "):
            p = doc.add_heading(stripped[3:], level=2)
            i += 1
            continue
        if stripped.startswith("### "):
            p = doc.add_heading(stripped[4:], level=3)
            i += 1
            continue

        # Table detection
        if "|" in stripped and i + 1 < len(lines) and re.match(r"^\|[-| :]+\|$", lines[i + 1].strip()):
            header_line = stripped
            i += 2  # skip separator
            table_rows = []
            while i < len(lines) and "|" in lines[i].strip() and lines[i].strip().startswith("|"):
                table_rows.append(lines[i])
                i += 1
            add_table_from_md(doc, header_line, table_rows)
            doc.add_paragraph("")  # spacer after table
            continue

        # Bold text line (like **Total this month:** $xxx)
        if stripped.startswith("**"):
            p = doc.add_paragraph()
            # Parse bold segments
            parts = re.split(r"(\*\*.*?\*\*)", stripped)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
            i += 1
            continue

        # Bullet points
        if stripped.startswith("• ") or stripped.startswith("- "):
            text = stripped[2:]
            p = doc.add_paragraph(text, style="List Bullet")
            i += 1
            continue

        # Sub-bullets
        if stripped.startswith("  - ") or stripped.startswith("  • "):
            text = stripped[4:]
            p = doc.add_paragraph(text, style="List Bullet 2")
            i += 1
            continue

        # Italic/generated line
        if stripped.startswith("_") and stripped.endswith("_"):
            p = doc.add_paragraph()
            run = p.add_run(stripped.strip("_"))
            run.italic = True
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(128, 128, 128)
            i += 1
            continue

        # Regular paragraph
        doc.add_paragraph(stripped)
        i += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    size_kb = output_path.stat().st_size / 1024
    print(f"Generated: {output_path} ({size_kb:.0f} KB)")


def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: generate-docx.py <markdown_file> <output_docx>", file=sys.stderr)
        sys.exit(1)

    md_path = pathlib.Path(sys.argv[1])
    out_path = pathlib.Path(sys.argv[2])

    if not md_path.exists():
        print(f"File not found: {md_path}", file=sys.stderr)
        sys.exit(1)

    content = md_path.read_text(encoding="utf-8").strip()
    if not content:
        print("Markdown file is empty — skipping.", file=sys.stderr)
        sys.exit(1)

    md_to_docx(content, out_path)


if __name__ == "__main__":
    main()
