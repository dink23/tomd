"""엑셀, 워드, 파워포인트, CSV 변환."""
import csv
import datetime
import decimal
import io
import re

from . import table
from .table import Cell


# ---------------------------------------------------------------- 엑셀

def _fmt(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
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


_BUILTIN_FORMATS = {
    "1": "0", "2": "0.00", "3": "#,##0", "4": "#,##0.00", "9": "0%", "10": "0.00%",
    "37": "#,##0 ;(#,##0)", "38": "#,##0 ;[Red](#,##0)", "39": "#,##0.00;(#,##0.00)", "40": "#,##0.00;[Red](#,##0.00)",
}


def _split_sections(code):
    parts, buf, quoted = [], "", False
    for ch in code:
        if ch == '"':
            quoted = not quoted
        if ch == ";" and not quoted:
            parts.append(buf)
            buf = ""
        else:
            buf += ch
    parts.append(buf)
    return parts


def _number_text(value, code):
    """엑셀 화면에 보이는 모양대로 숫자를 쓴다 (쉼표, 퍼센트, 소수 자릿수, △·명·원 같은 앞뒤 글자)."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return _fmt(value)
    code = (code or "").strip()
    if not code or code.lower() == "general" or code == "@":
        return _fmt(value)
    try:
        sections = _split_sections(code)
        section = sections[0]
        sign = ""
        if value < 0:
            if len(sections) > 1 and sections[1].strip():
                section = sections[1]  # 음수 전용 형식 (△1,000, (1,000) 등)
            else:
                sign = "-"
            value = -value
        elif value == 0 and len(sections) > 2:
            section = sections[2]  # 0 전용 형식 (비어 있으면 엑셀도 아무것도 보이지 않는다)
            if not section.strip():
                return ""
        tokens = re.findall(r'"[^"]*"|\\.|\[[^\]]*\]|_.|\*.|.', section)
        holders = [i for i, t in enumerate(tokens) if t in ("0", "#", "?")]

        def literal(items):
            out = ""
            for t in items:
                if t.startswith('"'):
                    out += t[1:-1]
                elif t.startswith("\\"):
                    out += t[1:]
                elif t.startswith("[$"):
                    out += t[2:-1].split("-", 1)[0]  # 통화 기호 ([$₩-412] → ₩)
                elif t.startswith(("[", "_", "*")):
                    continue
                else:
                    out += t
            return out

        if not holders:
            text = literal(tokens).strip()
            return text or _fmt(value)
        if any(t.lower() in ("e", "y", "m", "d", "h", "s", "/") for t in tokens):
            return _fmt(value)  # 지수·날짜·분수 형식은 값 그대로
        start, end = holders[0], holders[-1]
        while start > 0 and tokens[start - 1] in (".", ","):
            start -= 1
        scale = 0
        while end + 1 < len(tokens) and tokens[end + 1] == ",":
            end += 1
            scale += 1  # 형식 끝의 쉼표 하나마다 1000으로 나눠 보인다
        number = tokens[start:end + 1]
        prefix, suffix = literal(tokens[:start]), literal(tokens[end + 1:])
        body = "".join(number[: len(number) - scale])
        decimals = sum(1 for t in body.split(".", 1)[1] if t in "0#?") if "." in body else 0
        if "%" in prefix + suffix:
            value *= 100
        value /= 1000 ** scale
        grouped = "," in body
        # 엑셀처럼 반올림한다 (파이썬 기본은 짝수 쪽으로 반올림해 2.5가 2가 된다)
        rounded = decimal.Decimal(repr(value)).quantize(decimal.Decimal(1).scaleb(-decimals), rounding=decimal.ROUND_HALF_UP)
        text = f"{rounded:,f}" if grouped else f"{rounded:f}"
        if sign and float(text.replace(",", "")) == 0:
            sign = ""
        return f"{sign}{prefix}{text}{suffix}".strip()
    except Exception:
        return _fmt(value)


def _sheet_md(title, values, merges):
    """시트 하나를 마크다운으로. values: {(행, 열): 글}, merges: [(첫 행, 첫 열, 끝 행, 끝 열)]

    글이 하나도 없는 행과 열은 뺀다 (간격을 띄우려고 비워 둔 줄).
    """
    out = [f"## {title}"]
    if not values:
        out.append("(빈 시트)")
        return out
    rows = sorted({r for r, _ in values})
    cols = sorted({c for _, c in values})
    row_at = {r: i for i, r in enumerate(rows)}
    col_at = {c: i for i, c in enumerate(cols)}
    spans, covered = {}, set()
    for min_row, min_col, max_row, max_col in merges:
        if (min_row, min_col) not in values:
            continue  # 글 없는 병합은 서식일 뿐이다
        kept_r = [r for r in rows if min_row <= r <= max_row]
        kept_c = [c for c in cols if min_col <= c <= max_col]
        spans[(min_row, min_col)] = (len(kept_r), len(kept_c))
        for r in kept_r:
            for c in kept_c:
                if (r, c) != (min_row, min_col):
                    covered.add((r, c))
    cells = []
    for r in rows:
        for c in cols:
            if (r, c) in covered:
                continue
            rs, cs = spans.get((r, c), (1, 1))
            cells.append(Cell(row_at[r], col_at[c], values.get((r, c), ""), rs, cs))
    out.append(table.render(cells))
    return out


def _xlsx_openpyxl(path):
    import openpyxl
    from openpyxl.reader import excel as excel_reader

    # 차트·그림은 변환하지 않는다. 일부 차트에서 openpyxl이 읽다가 멈추므로 아예 읽지 않는다.
    excel_reader.find_images = lambda *args, **kwargs: ([], [])
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        # 실제로 있는 칸만 본다. 맨 아래 행·맨 오른쪽 열에 서식만 있는 파일에서
        # 사용 영역 전체(최대 100만 행)를 훑다가 멈추지 않도록.
        stored = getattr(ws, "_cells", None)
        if stored is None:
            stored = {(c.row, c.column): c for row in ws.iter_rows() for c in row}
        values = {}
        for key, cell in list(stored.items()):
            text = _number_text(cell.value, getattr(cell, "number_format", None))
            if text:
                values[key] = text
        merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
        sheets.append((ws.title, values, merges))
    return sheets


_DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))


def _is_date_format(code):
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', "", code or "").lower()
    if not code or code == "general":
        return False
    return bool(re.search(r"[ymdhs]", code)) and not re.search(r"[0#?]", code.split(";")[0].replace("0.0", ""))


def _xlsx_raw(path):
    """openpyxl이 못 읽는 파일을 위해 시트 XML에서 값만 직접 읽는다 (서식·차트 무시)."""
    import posixpath
    import zipfile
    from xml.etree import ElementTree as ET

    def local(tag):
        return tag.rsplit("}", 1)[-1]

    def attr(el, name):
        for key, value in el.attrib.items():
            if local(key) == name:
                return value
        return None

    z = zipfile.ZipFile(path)
    names = set(z.namelist())

    def rels_of(part):
        folder, base = posixpath.split(part)
        rel = posixpath.join(folder, "_rels", base + ".rels")
        out = {}
        if rel in names:
            for el in ET.fromstring(z.read(rel)):
                target = attr(el, "Target") or ""
                if attr(el, "TargetMode") == "External":
                    continue
                full = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
                out[attr(el, "Id")] = (attr(el, "Type") or "", full)
        return out

    book = next((p for t, p in rels_of("").values() if t.endswith("/officeDocument")), "xl/workbook.xml")
    book_rels = rels_of(book)
    root = ET.fromstring(z.read(book))
    date1904 = any(local(el.tag) == "workbookPr" and attr(el, "date1904") in ("1", "true") for el in root.iter())

    shared = []
    for kind, part in book_rels.values():
        if kind.endswith("/sharedStrings") and part in names:
            for si in ET.fromstring(z.read(part)):
                parts = []
                for el in si:
                    if local(el.tag) == "t":
                        parts.append(el.text or "")
                    elif local(el.tag) == "r":
                        parts.extend(t.text or "" for t in el if local(t.tag) == "t")
                shared.append("".join(parts))

    date_styles, style_codes = set(), {}
    for kind, part in book_rels.values():
        if kind.endswith("/styles") and part in names:
            styles = ET.fromstring(z.read(part))
            custom = {}
            xfs = []
            for el in styles:
                if local(el.tag) == "numFmts":
                    for f in el:
                        custom[attr(f, "numFmtId")] = attr(f, "formatCode")
                elif local(el.tag) == "cellXfs":
                    xfs = [attr(x, "numFmtId") or "0" for x in el]
            for i, fid in enumerate(xfs):
                if (fid.isdigit() and int(fid) in _DATE_IDS and fid not in custom) or _is_date_format(custom.get(fid)):
                    date_styles.add(str(i))
                else:
                    style_codes[str(i)] = custom.get(fid) or _BUILTIN_FORMATS.get(fid)

    def as_date(number):
        base = datetime.datetime(1904, 1, 1) if date1904 else datetime.datetime(1899, 12, 30)
        if number < 1:
            moment = base + datetime.timedelta(days=number)
            return moment.strftime("%H:%M")
        return base + datetime.timedelta(days=round(number * 86400) / 86400)

    def col_number(letters):
        n = 0
        for ch in letters:
            n = n * 26 + ord(ch) - 64
        return n

    sheets = []
    for el in root.iter():
        if local(el.tag) != "sheet" or attr(el, "state") in ("hidden", "veryHidden"):
            continue
        rid = attr(el, "id")
        if rid not in book_rels:
            continue
        part = book_rels[rid][1]
        if part not in names:
            continue
        values, merges = {}, []
        row, col = 0, 0
        for event, node in ET.iterparse(z.open(part), events=("start", "end")):
            tag = local(node.tag)
            if event == "start":
                if tag == "row":
                    row = int(attr(node, "r") or row + 1)
                    col = 0
                continue
            if tag == "c":
                ref = attr(node, "r")
                m = re.match(r"([A-Z]+)(\d+)", ref or "")
                if m:
                    col, row = col_number(m.group(1)), int(m.group(2))
                else:
                    col += 1
                kind = attr(node, "t")
                raw = None
                for child in node:
                    if local(child.tag) == "v":
                        raw = child.text
                    elif local(child.tag) == "is":
                        raw = "".join(t.text or "" for t in child.iter() if local(t.tag) == "t")
                if raw is not None and raw != "":
                    if kind == "s":
                        text = shared[int(raw)] if raw.isdigit() and int(raw) < len(shared) else ""
                    elif kind == "b":
                        text = "TRUE" if raw == "1" else "FALSE"
                    elif kind in ("str", "inlineStr", "e", "d"):
                        text = raw
                    else:
                        try:
                            number = float(raw)
                        except ValueError:
                            text = raw
                        else:
                            style = attr(node, "s")
                            if style in date_styles:
                                text = _fmt(as_date(number))
                            else:
                                if number.is_integer() and abs(number) < 2 ** 53:
                                    number = int(number)
                                text = _number_text(number, style_codes.get(style))
                    if text:
                        values[(row, col)] = text
                node.clear()
            elif tag == "mergeCell":
                m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", attr(node, "ref") or "")
                if m:
                    merges.append((int(m.group(2)), col_number(m.group(1)), int(m.group(4)), col_number(m.group(3))))
            elif tag == "row":
                node.clear()
        sheets.append((attr(el, "name") or "", values, merges))
    return sheets


def _xlsx_drawings(path):
    """시트 이름별 (차트 수, 그림 수). 차트·그림은 변환하지 않고 있다는 것만 알린다."""
    import posixpath
    import zipfile
    from xml.etree import ElementTree as ET

    def local(tag):
        return tag.rsplit("}", 1)[-1]

    def attr(el, name):
        for key, value in el.attrib.items():
            if local(key) == name:
                return value
        return None

    found = {}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        def rels_of(part):
            folder, base = posixpath.split(part)
            rel = posixpath.join(folder, "_rels", base + ".rels")
            out = {}
            if rel in names:
                for el in ET.fromstring(z.read(rel)):
                    target = attr(el, "Target") or ""
                    if attr(el, "TargetMode") == "External":
                        continue
                    full = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
                    out[attr(el, "Id")] = (attr(el, "Type") or "", full)
            return out

        book = next((p for t, p in rels_of("").values() if t.endswith("/officeDocument")), "xl/workbook.xml")
        book_rels = rels_of(book)
        for el in ET.fromstring(z.read(book)).iter():
            if local(el.tag) != "sheet" or attr(el, "id") not in book_rels:
                continue
            charts = pictures = 0
            for kind, drawing in rels_of(book_rels[attr(el, "id")][1]).values():
                if not kind.endswith("/drawing") or drawing not in names:
                    continue
                charts += sum(1 for t, _ in rels_of(drawing).values() if t.endswith("/chart"))
                pictures += sum(1 for n in ET.fromstring(z.read(drawing)).iter() if local(n.tag) == "pic")
            found[attr(el, "name") or ""] = (charts, pictures)
    return found


def _head(path, size=512):
    with open(path, "rb") as f:
        return f.read(size)


def convert_xlsx(path):
    head = _head(path)
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        import olefile

        if olefile.isOleFile(path):
            with olefile.OleFileIO(path) as ole:
                if ole.exists("EncryptedPackage"):
                    raise ValueError("암호가 걸린 엑셀 파일입니다. 엑셀에서 암호를 해제해 저장한 뒤 다시 변환하세요.")
        return convert_xls(path)  # 확장자만 xlsx인 예전 형식(xls)
    try:
        sheets = _xlsx_openpyxl(path)
    except Exception as first:
        try:
            sheets = _xlsx_raw(path)
        except Exception:
            raise first
    try:
        drawings = _xlsx_drawings(path)
    except Exception:
        drawings = {}
    out = []
    for title, values, merges in sheets:
        out.extend(_sheet_md(title, values, merges))
        charts, pictures = drawings.get(title, (0, 0))
        notes = [f"차트 {charts}개" if charts else "", f"그림 {pictures}개" if pictures else ""]
        if any(notes):
            out.append("[" + ", ".join(n for n in notes if n) + "]")
    return "\n\n".join(out)


def convert_xls(path):
    head = _head(path, 2048)
    if head.startswith(b"PK"):
        return convert_xlsx(path)  # 확장자만 xls인 새 형식(xlsx)
    if not head.startswith(b"\xd0\xcf\x11\xe0"):
        # 나이스·에듀파인 등에서 내려받은 xls는 실제로는 HTML이나 XML 표인 경우가 많다
        text = read_text(path)
        if re.search(r"urn:schemas-microsoft-com:office:spreadsheet", text[:5000]):
            return _spreadsheet_xml(text)
        if re.search(r"<table", text, re.I):
            return _html_tables(text)
        if "\t" in text or "," in text:
            # 확장자만 xls인 텍스트 표 (탭·쉼표로 구분)
            delimiter = "\t" if text.count("\t") >= text.count(",") else ","
            rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
            width = max((len(r) for r in rows), default=0)
            return table.render(table.from_grid([r + [""] * (width - len(r)) for r in rows]))
    import xlrd

    def open_book(**options):
        try:
            return xlrd.open_workbook(path, **options)
        except xlrd.biffh.XLRDError as exc:
            if "encrypted" in str(exc).lower():
                raise ValueError("암호가 걸린 엑셀 파일입니다. 엑셀에서 암호를 해제해 저장한 뒤 다시 변환하세요.") from exc
            raise

    try:
        book, formatted = open_book(formatting_info=True), True  # 병합·숫자 모양을 읽으려면 필요하다
    except ValueError:
        raise
    except Exception:
        book, formatted = open_book(), False
    out = []
    for index, sheet in enumerate(book.sheets()):
        if getattr(sheet, "visibility", 0):
            continue  # 숨긴 시트
        values = {}
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                value = cell.value
                code = None
                if formatted:
                    try:
                        code = book.format_map[book.xf_list[cell.xf_index].format_key].format_str
                    except Exception:
                        code = None
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(value, book.datemode)
                    except Exception:
                        pass
                    text = _fmt(value)
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    if float(value).is_integer() and abs(value) < 2 ** 53:
                        value = int(value)
                    text = _number_text(value, code)
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    text = "TRUE" if value else "FALSE"
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    text = xlrd.error_text_from_code.get(value, "")
                else:
                    text = _fmt(value)
                if text:
                    values[(r + 1, c + 1)] = text
        merges = [(r0 + 1, c0 + 1, r1, c1) for r0, r1, c0, c1 in (sheet.merged_cells if formatted else [])]
        out.extend(_sheet_md(sheet.name, values, merges))
    return "\n\n".join(out)


def _grid_cells(rows):
    """[(글, 행 병합 수, 열 병합 수)] 행 목록을 병합을 고려한 Cell 목록으로."""
    taken, cells = set(), []
    for r, row in enumerate(rows):
        c = 0
        for text, rs, cs in row:
            while (r, c) in taken:
                c += 1
            for dr in range(rs):
                for dc in range(cs):
                    taken.add((r + dr, c + dc))
            cells.append(Cell(r, c, text, rs, cs))
            c += cs
    return cells


def _html_tables(text):
    from html.parser import HTMLParser

    class Parser(HTMLParser):
        """표마다 행·칸을 모은다. 표 안의 표는 바깥 칸의 글로 넣는다."""

        def __init__(self):
            super().__init__(convert_charrefs=True)
            self.tables, self.stack = [], []  # stack: [{"rows": [...], "cell": [글 조각, 행 병합, 열 병합] 또는 None}]

        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            if tag == "table":
                self.stack.append({"rows": [], "cell": None})
            elif not self.stack:
                return
            elif tag == "tr":
                self.stack[-1]["rows"].append([])
            elif tag in ("td", "th"):
                frame = self.stack[-1]
                if not frame["rows"]:
                    frame["rows"].append([])

                def num(key):
                    try:
                        return max(1, min(1000, int(a.get(key) or 1)))
                    except ValueError:
                        return 1
                frame["cell"] = [[], num("rowspan"), num("colspan")]
            elif tag == "br" and self.stack[-1]["cell"] is not None:
                self.stack[-1]["cell"][0].append("\n")

        def handle_endtag(self, tag):
            if not self.stack:
                return
            frame = self.stack[-1]
            if tag in ("td", "th") and frame["cell"] is not None:
                text = re.sub(r"[ \t\r\f\v]+", " ", "".join(frame["cell"][0]))
                text = "\n".join(t.strip() for t in text.split("\n") if t.strip())
                frame["rows"][-1].append((text, frame["cell"][1], frame["cell"][2]))
                frame["cell"] = None
            elif tag == "table":
                self.stack.pop()
                rows = [r for r in frame["rows"] if r]
                if self.stack and self.stack[-1]["cell"] is not None:
                    inner = "\n".join(" / ".join(t for t, _, _ in r if t) for r in rows)
                    self.stack[-1]["cell"][0].append("\n" + inner + "\n")
                elif rows:
                    self.tables.append(rows)

        def handle_data(self, data):
            if self.stack and self.stack[-1]["cell"] is not None:
                self.stack[-1]["cell"][0].append(data)

    p = Parser()
    p.feed(text)
    p.close()
    out = [table.render(_grid_cells(rows)) for rows in p.tables]
    return "\n\n".join(o for o in out if o) or "(빈 표)"


def _spreadsheet_xml(text):
    from xml.etree import ElementTree as ET

    def local(tag):
        return tag.rsplit("}", 1)[-1]

    def attr(el, name):
        for key, value in el.attrib.items():
            if local(key) == name:
                return value
        return None

    root = ET.fromstring(text.encode("utf-8") if not text.lstrip().startswith("<?xml") else re.sub(r"^\s*<\?xml[^>]*\?>", "", text))
    out = []
    for ws in (el for el in root.iter() if local(el.tag) == "Worksheet"):
        values, merges = {}, []
        r = 0
        for row in (el for el in ws.iter() if local(el.tag) == "Row"):
            r = int(attr(row, "Index") or r + 1)
            c = 0
            for cell in (el for el in row if local(el.tag) == "Cell"):
                c = int(attr(cell, "Index") or c + 1)
                data = "".join("".join(d.itertext()) for d in cell if local(d.tag) == "Data").strip()
                if data:
                    values[(r, c)] = data
                across, down = int(attr(cell, "MergeAcross") or 0), int(attr(cell, "MergeDown") or 0)
                if across or down:
                    merges.append((r, c, r + down, c + across))
                c += across
        out.extend(_sheet_md(attr(ws, "Name") or "", values, merges))
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
    return table.render(cells, single="quote")


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


def _guess_title(slide, slide_height):
    """제목 개체 틀 없이 글상자로만 만든 슬라이드에서 제목으로 쓸 글상자를 고른다.

    슬라이드 위쪽에 있는 짧은 글상자 가운데 글자가 가장 큰 것(같으면 더 위의 것)을 고른다.
    """
    found = []

    def walk(shapes):
        for shape in shapes:
            try:
                if shape.shape_type == 6:
                    walk(shape.shapes)
                    continue
                if not shape.has_text_frame or _ph_type(shape):
                    continue
                paras = [p for p in shape.text_frame.paragraphs if _para_text(p)]
                text = " ".join(_para_text(p) for p in paras)
                if not paras or len(paras) > 2 or len(text) > 60 or _BULLET.match(text):
                    continue
                top = shape.top
                if top is None or not slide_height or top > 0.3 * slide_height:
                    continue
                sizes = [r.font.size.pt for p in paras for r in p.runs if r.font.size]
                found.append((-(max(sizes) if sizes else 0), top, len(found), shape))
            except Exception:
                continue

    walk(slide.shapes)
    return min(found)[3] if found else None


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
        if title_shape is None:
            title_shape = _guess_title(slide, prs.slide_height)
        title = " ".join(title_shape.text_frame.text.split()) if title_shape is not None else ""
        hidden = slide._element.get("show") in ("0", "false")
        heading = f"## 슬라이드 {n}" + (" (숨김)" if hidden else "") + (f": {title}" if title else "")
        out.append(heading)
        pictures = [0]

        def walk(shapes):
            for shape in shapes:
                # 도형 번호(id)는 복사·붙여넣기한 파일에서 겹칠 수 있으므로 문서 속 요소 자체로 비교한다
                if title_shape is not None and shape._element is title_shape._element:
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
