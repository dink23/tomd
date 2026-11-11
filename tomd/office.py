"""엑셀, 워드, 파워포인트, CSV 변환."""
import csv
import datetime
import io
import re

from . import table
from .table import Cell


# ---------------------------------------------------------------- 엑셀

def _fmt(value):
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.10g}"
    if isinstance(value, datetime.datetime):
        if value.time() == datetime.time(0, 0):
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, datetime.date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def convert_xlsx(path):
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        out.append(f"## {ws.title}")
        values = {}
        for row in ws.iter_rows():
            for cell in row:
                text = _fmt(cell.value)
                if text:
                    values[(cell.row, cell.column)] = text
        if not values:
            out.append("(빈 시트)")
            continue
        r0 = min(r for r, _ in values)
        r1 = max(r for r, _ in values)
        c0 = min(c for _, c in values)
        c1 = max(c for _, c in values)
        spans, covered = {}, set()
        for rng in ws.merged_cells.ranges:
            spans[(rng.min_row, rng.min_col)] = (
                rng.max_row - rng.min_row + 1,
                rng.max_col - rng.min_col + 1,
            )
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        covered.add((r, c))
        cells = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) in covered:
                    continue
                rs, cs = spans.get((r, c), (1, 1))
                rs = min(rs, r1 - r + 1)
                cs = min(cs, c1 - c + 1)
                cells.append(Cell(r - r0, c - c0, values.get((r, c), ""), rs, cs))
        out.append(table.render(cells))
    return "\n\n".join(out)


def convert_xls(path):
    import xlrd

    book = xlrd.open_workbook(path)
    out = []
    for sheet in book.sheets():
        out.append(f"## {sheet.name}")
        grid = []
        for r in range(sheet.nrows):
            row = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, book.datemode)
                    except Exception:
                        pass
                row.append(_fmt(value))
            grid.append(row)
        rendered = table.render(table.from_grid(grid)) if grid else ""
        out.append(rendered or "(빈 시트)")
    return "\n\n".join(out)


def read_text(path):
    data = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp949", "utf-16"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def convert_csv(path):
    text = read_text(path)
    delimiter = "\t" if str(path).lower().endswith(".tsv") else ","
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    width = max((len(r) for r in rows), default=0)
    grid = [r + [""] * (width - len(r)) for r in rows]
    return table.render(table.from_grid(grid))


# ---------------------------------------------------------------- 워드

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_runs(p):
    """문단의 글자를 굵게/기울임 표시와 함께 이어 붙인다."""
    pieces = []  # (text, bold, italic)
    for r in p.iter(W + "r"):
        text = ""
        for node in r:
            if node.tag == W + "t":
                text += node.text or ""
            elif node.tag == W + "tab":
                text += "\t"
            elif node.tag in (W + "br", W + "cr"):
                text += "\n"
        if not text:
            continue
        rpr = r.find(W + "rPr")
        bold = italic = False
        if rpr is not None:
            b = rpr.find(W + "b")
            i = rpr.find(W + "i")
            bold = b is not None and b.get(W + "val") not in ("0", "false")
            italic = i is not None and i.get(W + "val") not in ("0", "false")
        if pieces and pieces[-1][1:] == (bold, italic):
            pieces[-1] = (pieces[-1][0] + text, bold, italic)
        else:
            pieces.append((text, bold, italic))
    return pieces


def _docx_text(p, markup=True):
    out = ""
    for text, bold, italic in _docx_runs(p):
        core = text.strip()
        if markup and core and (bold or italic) and "\n" not in core:
            mark = "**" if bold else ""
            mark += "*" if italic else ""
            lead = text[: len(text) - len(text.lstrip())]
            tail = text[len(text.rstrip()):]
            out += f"{lead}{mark}{core}{mark[::-1]}{tail}"
        else:
            out += text
    return out.strip()


def _docx_table(tbl):
    cells = []
    open_merge = {}  # col -> Cell (세로 병합 진행 중)
    for r, tr in enumerate(tbl.findall(W + "tr")):
        col = 0
        for tc in tr.findall(W + "tc"):
            pr = tc.find(W + "tcPr")
            span, vmerge = 1, None
            if pr is not None:
                gs = pr.find(W + "gridSpan")
                if gs is not None:
                    span = int(gs.get(W + "val", "1"))
                vm = pr.find(W + "vMerge")
                if vm is not None:
                    vmerge = vm.get(W + "val", "continue")
            if vmerge == "continue" and col in open_merge:
                open_merge[col].rowspan += 1
            else:
                text = "\n".join(
                    t for t in (_docx_text(p, markup=False) for p in tc.iter(W + "p")) if t
                )
                cell = Cell(r, col, text, 1, span)
                cells.append(cell)
                if vmerge == "restart":
                    open_merge[col] = cell
                else:
                    open_merge.pop(col, None)
            col += span
    return table.render(cells)


def convert_docx(path):
    import docx

    doc = docx.Document(path)
    styles = {}
    for style in doc.styles:
        try:
            styles[style.style_id] = style.name or ""
        except Exception:
            pass
    out = []
    for el in doc.element.body:
        if el.tag == W + "tbl":
            rendered = _docx_table(el)
            if rendered:
                out.append(rendered)
            continue
        if el.tag != W + "p":
            continue
        ppr = el.find(W + "pPr")
        style_name, is_list = "", False
        if ppr is not None:
            ps = ppr.find(W + "pStyle")
            if ps is not None:
                style_name = styles.get(ps.get(W + "val"), ps.get(W + "val") or "")
            is_list = ppr.find(W + "numPr") is not None
        low = style_name.lower()
        m = re.match(r"(heading|제목)\s*(\d)", low)
        if m or low == "title":
            text = _docx_text(el, markup=False)
            if text:
                level = 1 if not m else min(int(m.group(2)), 6)
                out.append("#" * level + " " + text.replace("\n", " "))
            continue
        text = _docx_text(el)
        if not text:
            continue
        if is_list or low.startswith("list"):
            marker = "1. " if "number" in low else "- "
            item = marker + text.replace("\n", " ")
            if out and re.match(r"(- |1\. )", out[-1].splitlines()[-1]):
                out[-1] += "\n" + item
            else:
                out.append(item)
        else:
            out.append(text.replace("\n", "  \n"))
    return "\n\n".join(out)


# ---------------------------------------------------------------- 파워포인트

def convert_pptx(path):
    from pptx import Presentation

    prs = Presentation(path)
    out = []
    for n, slide in enumerate(prs.slides, 1):
        title_shape = slide.shapes.title
        title = title_shape.text_frame.text.strip() if title_shape is not None else ""
        out.append(f"## 슬라이드 {n}" + (f": {' '.join(title.split())}" if title else ""))

        def walk(shapes):
            for shape in shapes:
                if title_shape is not None and shape.shape_id == title_shape.shape_id:
                    continue
                if shape.shape_type == 6:  # 그룹
                    walk(shape.shapes)
                elif getattr(shape, "has_table", False) and shape.has_table:
                    grid = [[c.text for c in row.cells] for row in shape.table.rows]
                    rendered = table.render(table.from_grid(grid))
                    if rendered:
                        out.append(rendered)
                elif shape.has_text_frame:
                    lines = []
                    for para in shape.text_frame.paragraphs:
                        text = "".join(run.text for run in para.runs).strip()
                        if text:
                            lines.append("  " * para.level + "- " + text)
                    if lines:
                        out.append("\n".join(lines))

        walk(slide.shapes)
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                out.append("> 발표자 노트: " + notes.replace("\n", "\n> "))
    return "\n\n".join(out)
