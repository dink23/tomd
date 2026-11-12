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

_BULLET = re.compile(r"^\s*[•·▪■□◦○●∙‣◆◇▶▷►\-–—]\s+")
_P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def _ph_type(shape):
    """자리 표시자 종류 이름(TITLE, BODY, SLIDE_NUMBER …). 자리 표시자가 아니면 빈 문자열."""
    try:
        if shape.is_placeholder:
            return str(shape.placeholder_format.type).split(".")[-1].split(" ")[0]
    except Exception:
        pass
    return ""


def _para_text(para):
    # 문단 안의 줄바꿈(Shift+Enter)은 python-pptx가 세로 탭으로 돌려준다.
    # 대개 글상자 너비에 맞춘 줄 나눔이므로 빈칸으로 잇는다.
    return re.sub(r" {2,}", " ", (para.text or "").replace("\x0b", " ")).strip()


def _text_frame_md(shape):
    """글상자 하나를 마크다운으로. 목록이면 글머리표, 아니면 줄글로 만든다."""
    paras = [(p.level or 0, _para_text(p)) for p in shape.text_frame.paragraphs]
    paras = [(lv, t) for lv, t in paras if t]
    if not paras:
        return ""
    body = _ph_type(shape) in ("BODY", "OBJECT", "SUBTITLE") and len(paras) > 1
    marked = any(_BULLET.match(t) for _, t in paras)
    if not body and not marked:
        # 목록이 아닌 글상자: 문단마다 줄을 바꿔 한 덩어리로
        return "  \n".join(t.replace("\n", "  \n") for _, t in paras)
    lines, seen_marker = [], False
    for level, text in paras:
        flat = " ".join(text.split("\n"))
        if marked:
            if _BULLET.match(flat):
                seen_marker = True
                lines.append("  " * level + "- " + _BULLET.sub("", flat))
            elif seen_marker:
                # 글머리표 항목 바로 아래의 설명 줄은 한 단계 들여 쓴다
                lines.append("  " * (level + 1) + "- " + flat)
            else:
                lines.append(flat + "  ")
        else:
            lines.append("  " * level + "- " + flat)
    return "\n".join(lines).rstrip()


def convert_pptx(path):
    from pptx import Presentation

    prs = Presentation(path)
    out = []
    for n, slide in enumerate(prs.slides, 1):
        title_shape = None
        for shape in slide.shapes:
            if _ph_type(shape) in ("TITLE", "CENTER_TITLE") and shape.has_text_frame:
                if shape.text_frame.text.strip():
                    title_shape = shape
                    break
        title = " ".join(title_shape.text_frame.text.split()) if title_shape is not None else ""
        hidden = slide._element.get("show") in ("0", "false")
        heading = f"## 슬라이드 {n}" + (" (숨김)" if hidden else "") + (f": {title}" if title else "")
        out.append(heading)
        pictures = [0]

        def walk(shapes):
            for shape in shapes:
                if title_shape is not None and shape.shape_id == title_shape.shape_id:
                    continue
                kind = _ph_type(shape)
                if kind in ("SLIDE_NUMBER", "DATE"):
                    continue  # 쪽 번호·날짜 칸
                if shape.shape_type == 6:  # 그룹
                    walk(shape.shapes)
                elif shape.shape_type == 13:  # 그림
                    pictures[0] += 1
                elif getattr(shape, "has_table", False) and shape.has_table:
                    grid = [[c.text for c in row.cells] for row in shape.table.rows]
                    rendered = table.render(table.from_grid(grid))
                    if rendered:
                        out.append(rendered)
                elif shape.has_text_frame:
                    rendered = _text_frame_md(shape)
                    if rendered:
                        out.append(rendered)

        walk(slide.shapes)
        if pictures[0]:
            out.append(f"[그림 {pictures[0]}개]")
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                out.append("> 발표자 노트: " + notes.replace("\x0b", "\n").replace("\n", "\n> "))
    return "\n\n".join(out)
