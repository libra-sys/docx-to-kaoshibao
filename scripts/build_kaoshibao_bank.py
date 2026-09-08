"""
Generic utility: generate a Kaoshibao-compatible Excel question bank from a list of questions.

Usage:
    python build_kaoshibao_bank.py <input_json> <output_xlsx>

Input JSON format (list of objects):
[
    {
        "stem": "If water ______ heated, it will be ______ into vapour.",
        "A": "was, got",
        "B": "is, get",
        "C": "is, turned",
        "D": "was, turned",
        "answer": "C",
        "translation": "如果水被加热，它会变成蒸汽。",
        "explanation": "在条件状语从句中，用一般现在时表将来..."
    },
    ...
]

The output Excel uses the standard Kaoshibao import column headers:
    题型, 题干, 选项A, 选项B, 选项C, 选项D, 正确答案, 解析, 难度
"""

import json
import sys
from pathlib import Path

def generate_excel(questions, output_path):
    try:
        import openpyxl
    except ImportError:
        print("Error: openpyxl is not installed. Install it with: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "题库"

    headers = ["题型", "题干", "选项A", "选项B", "选项C", "选项D", "正确答案", "解析", "难度"]
    ws.append(headers)

    # Header style
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for idx, q in enumerate(questions, start=2):
        explanation = f"【句意】{q.get('translation', '')}\n【解析】{q.get('explanation', '')}"
        row = [
            "单选题",
            q.get("stem", ""),
            q.get("A", ""),
            q.get("B", ""),
            q.get("C", ""),
            q.get("D", ""),
            q.get("answer", ""),
            explanation,
            q.get("difficulty", "中等")
        ]
        ws.append(row)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=idx, column=col)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = thin_border
            if col == 7:
                cell.font = Font(bold=True, color="C00000")

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 60
    ws.column_dimensions["I"].width = 10

    ws.row_dimensions[1].height = 25
    for idx in range(2, len(questions) + 2):
        ws.row_dimensions[idx].height = 80

    wb.save(output_path)
    print(f"Saved {len(questions)} questions to: {output_path}")


def extract_docx_text(docx_path):
    """Extract plain text paragraphs from a .docx file without external dependencies."""
    import zipfile
    import xml.etree.ElementTree as ET

    paragraphs = []
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for p in root.findall('.//w:p', ns):
                texts = []
                for r in p.findall('.//w:r', ns):
                    for t in r.findall('.//w:t', ns):
                        if t.text:
                            texts.append(t.text)
                line = ''.join(texts)
                if line.strip():
                    paragraphs.append(line)
    return paragraphs


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python build_kaoshibao_bank.py <input_json> <output_xlsx>")
        sys.exit(1)

    input_json = Path(sys.argv[1])
    output_xlsx = Path(sys.argv[2])

    with open(input_json, 'r', encoding='utf-8') as f:
        questions = json.load(f)

    generate_excel(questions, output_xlsx)
