"""엑셀, 워드, 파워포인트, CSV 변환."""
import csv
import datetime
import decimal
import io
import os
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


_EPOCH_1900 = datetime.datetime(1899, 12, 30)
_EPOCH_1904 = datetime.datetime(1904, 1, 1)
# 엑셀 기본 날짜 형식(번호 14·22)은 화면에서 지역 설정을 따른다: 한국어 윈도에서는 2027-03-02
_LOCALE_DATES = {"mm-dd-yy": "yyyy-mm-dd", "m/d/yy": "yyyy-mm-dd", "m/d/yyyy": "yyyy-mm-dd",
                 "m/d/yy h:mm": "yyyy-mm-dd h:mm", "m/d/yy hh:mm": "yyyy-mm-dd hh:mm"}
_DATE_TOKEN = re.compile(
    r'"[^"]*"|\\.|\[[^\]]*\]|_.|\*.|yyyy|yy|y|e|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|'
    r'hh|h|ss|s|am/pm|a/p|上午/下午|오전/오후|\.0+|.',
    re.I,
)
_WEEK_KO = "월화수목금토일"
_WEEK_EN = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
_MONTH_EN = ("January", "February", "March", "April", "May", "June", "July", "August",
             "September", "October", "November", "December")


def _serial(value, date1904=False):
    """날짜·시각 값을 엑셀 일련번호(일 단위 수)로."""
    base = _EPOCH_1904 if date1904 else _EPOCH_1900
    if isinstance(value, datetime.datetime):
        return (value - base).total_seconds() / 86400
    if isinstance(value, datetime.date):
        return (datetime.datetime(value.year, value.month, value.day) - base).days
    if isinstance(value, datetime.time):
        return (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6) / 86400
    if isinstance(value, datetime.timedelta):
        return value.total_seconds() / 86400
    return float(value)


def _is_date_code(code):
    return _is_date_format(code) and not re.search(r"E[+-]", code or "")


def _date_text(value, code, date1904=False):
    """날짜·시각을 엑셀 서식 그대로 (yyyy"년" m"월" d"일", m/d, hh:mm, [h]:mm, 오전/오후 …)."""
    code = _LOCALE_DATES.get((code or "").lower(), code)
    serial = _serial(value, date1904)
    section = _split_sections(code)[0]
    base = _EPOCH_1904 if date1904 else _EPOCH_1900
    seconds = round(serial * 86400)
    moment = base + datetime.timedelta(seconds=seconds)
    tokens = _DATE_TOKEN.findall(section)
    lowered = [t.lower() for t in tokens]
    korean = bool(re.search(r"\[\$-(?:0*412|ko)", section, re.I))  # 한국어 지역 서식이면 요일도 한국어
    ampm = any(t in ("am/pm", "a/p", "오전/오후", "上午/下午") for t in lowered)
    kinds = []  # 날짜·시각 자리표시자만 (분/월 구분용)
    for i, t in enumerate(lowered):
        if t[:1] in "ymdhs" and not t.startswith(("[", '"')) or re.fullmatch(r"\[(h+|m+|s+)\]", t):
            kinds.append(i)
    out = []
    for i, (t, low) in enumerate(zip(tokens, lowered)):
        if low.startswith('"'):
            out.append(t[1:-1])
        elif low.startswith("\\"):
            out.append(t[1:])
        elif low.startswith("["):
            elapsed = re.fullmatch(r"\[(h+|m+|s+)\]", low)
            if elapsed:
                unit = elapsed.group(1)
                total = {"h": seconds // 3600, "m": seconds // 60, "s": seconds}[unit[0]]
                out.append(str(int(total)).zfill(len(unit)))
            elif low.startswith("[$") and "-" in low:
                out.append(t[2:-1].split("-", 1)[0])
            # [Red] 같은 색·조건은 보이지 않는다
        elif low.startswith(("_", "*")):
            continue
        elif low in ("yyyy", "e"):
            out.append(f"{moment.year:04d}")
        elif low in ("yy", "y"):
            out.append(f"{moment.year % 100:02d}")
        elif low[:1] == "m" and low.strip("m") == "":
            pos = kinds.index(i) if i in kinds else -1
            prev_k = lowered[kinds[pos - 1]] if pos > 0 else ""
            next_k = lowered[kinds[pos + 1]] if 0 <= pos < len(kinds) - 1 else ""
            minute = len(low) <= 2 and (prev_k.lstrip("[")[:1] == "h" or next_k[:1] == "s")
            if minute:
                out.append(f"{moment.minute:0{len(low)}d}")
            elif len(low) == 1:
                out.append(str(moment.month))
            elif len(low) == 2:
                out.append(f"{moment.month:02d}")
            elif len(low) == 3:
                out.append(_MONTH_EN[moment.month - 1][:3])
            elif len(low) == 4:
                out.append(_MONTH_EN[moment.month - 1])
            else:
                out.append(_MONTH_EN[moment.month - 1][0])
        elif low == "d":
            out.append(str(moment.day))
        elif low == "dd":
            out.append(f"{moment.day:02d}")
        elif low == "ddd":
            out.append(_WEEK_KO[moment.weekday()] if korean else _WEEK_EN[moment.weekday()][:3])
        elif low == "dddd":
            out.append(_WEEK_KO[moment.weekday()] + "요일" if korean else _WEEK_EN[moment.weekday()])
        elif low == "aaa":
            out.append(_WEEK_KO[moment.weekday()])
        elif low == "aaaa":
            out.append(_WEEK_KO[moment.weekday()] + "요일")
        elif low in ("h", "hh"):
            hour = moment.hour
            if ampm:
                hour = hour % 12 or 12
            out.append(f"{hour:0{len(low)}d}")
        elif low in ("s", "ss"):
            out.append(f"{moment.second:0{len(low)}d}")
        elif low.startswith(".0"):
            frac = serial * 86400 - int(serial * 86400)
            out.append("." + f"{frac:.{len(low) - 1}f}"[2:])
        elif low in ("am/pm", "a/p"):
            pm = moment.hour >= 12
            word = ("PM" if pm else "AM") if low == "am/pm" else ("P" if pm else "A")
            out.append(word if t[0].isupper() else word.lower())
        elif low == "오전/오후":
            out.append("오후" if moment.hour >= 12 else "오전")
        elif low == "上午/下午":
            out.append("下午" if moment.hour >= 12 else "上午")
        else:
            out.append(t)
    return "".join(out).strip()


def _fraction_text(value, section):
    """# ?/? 같은 분수 서식."""
    from fractions import Fraction

    body = re.sub(r'"[^"]*"|\[[^\]]*\]|_.|\*.', "", section).strip()
    m = re.match(r"^([#0?,]*)\s*([#0?]+)\s*/\s*([#0?]+|\d+)$", body)
    if not m:
        return None
    whole_fmt, den_fmt = m.group(1), m.group(3)
    sign = "-" if value < 0 else ""
    value = abs(value)
    whole = int(value) if whole_fmt else 0
    rest = value - whole
    if den_fmt.isdigit():
        den = int(den_fmt)
        num = round(rest * den)
    else:
        frac = Fraction(rest).limit_denominator(10 ** len(den_fmt) - 1)
        num, den = frac.numerator, frac.denominator
    if num == den:
        whole, num = whole + 1, 0
    if num == 0:
        return f"{sign}{whole}"
    head = f"{whole} " if whole else ""
    return f"{sign}{head}{num}/{den}".strip()


def _exponent_text(value, section):
    """0.00E+00 같은 지수 서식."""
    body = re.sub(r'"[^"]*"|\[[^\]]*\]|_.|\*.', "", section).strip()
    m = re.match(r"^([#0]*)(?:\.([#0]*))?E([+-])([#0]+)$", body, re.I)
    if not m:
        return None
    decimals = len(m.group(2) or "")
    digits = len(m.group(4))
    mantissa, exponent = f"{value:.{decimals}E}".split("E")
    exp = int(exponent)
    sign = "-" if exp < 0 else ("+" if m.group(3) == "+" else "")
    return f"{mantissa}E{sign}{abs(exp):0{digits}d}"


def _cell_text(value, code, date1904=False):
    """칸 값을 엑셀 화면에 보이는 모양대로."""
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)):
        if code and _is_date_code(code):
            try:
                return _date_text(value, code, date1904)
            except Exception:
                pass
        if isinstance(value, datetime.time):
            return value.strftime("%H:%M") if not value.second else value.strftime("%H:%M:%S")
        if isinstance(value, datetime.timedelta):
            return str(value)
        return _fmt(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool) and code and _is_date_code(code):
        try:
            return _date_text(value, code, date1904)
        except Exception:
            return _fmt(value)
    return _number_text(value, code)


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
        if "/" in tokens:
            text = _fraction_text(-value if sign else value, section)
            return text if text is not None else _fmt(-value if sign else value)
        if any(t.lower() == "e" for t in tokens):
            text = _exponent_text(value, section)
            return (sign + text) if text is not None else _fmt(-value if sign else value)
        if any(t.lower() in ("y", "m", "d", "h", "s") for t in tokens):
            return _fmt(-value if sign else value)
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


def _link_text(text, target=None, location=None):
    """칸에 걸린 링크: 바깥 주소는 [글](주소), 문서 안 위치는 글 (→ 시트!칸)."""
    if target:
        return f"[{text}]({target.replace(' ', '%20').replace(')', '%29')})"
    if location:
        return f"{text} (→ {location.replace(chr(39), '')})"
    return text


def _hidden_note(rows, cols):
    parts = []
    if rows:
        parts.append(f"숨긴 행 {rows}개")
    if cols:
        parts.append(f"숨긴 열 {cols}개")
    return f"> {'와 '.join(parts)}는 변환하지 않았습니다 (엑셀 화면에 보이는 내용만 옮김)." if parts else ""


def _drop_hidden(values, merges, hidden_rows, hidden_cols):
    """숨긴 행·열의 칸을 뺀다. 글이 있던 숨긴 행·열 수를 함께 돌려준다."""
    gone_r = {r for r, c in values if r in hidden_rows}
    gone_c = {c for r, c in values if c in hidden_cols}
    if not gone_r and not gone_c:
        return values, merges, "", 0
    kept = {k: v for k, v in values.items() if k[0] not in hidden_rows and k[1] not in hidden_cols}
    return kept, merges, _hidden_note(len(gone_r), len(gone_c)), 1


def _hidden_sheets_note(names):
    if not names:
        return ""
    return f"> 숨긴 시트 {len(names)}개({', '.join(names)})는 변환하지 않았습니다."


def _xlsx_openpyxl(path):
    import openpyxl
    from openpyxl.reader import excel as excel_reader
    from openpyxl.utils import column_index_from_string

    # 차트·그림은 변환하지 않는다. 일부 차트에서 openpyxl이 읽다가 멈추므로 아예 읽지 않는다.
    excel_reader.find_images = lambda *args, **kwargs: ([], [])
    wb = openpyxl.load_workbook(path, data_only=True)
    date1904 = bool(getattr(getattr(wb, "epoch", None), "year", 1900) == 1904)
    sheets, hidden_sheets = [], []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            hidden_sheets.append(ws.title)
            continue
        # 실제로 있는 칸만 본다. 맨 아래 행·맨 오른쪽 열에 서식만 있는 파일에서
        # 사용 영역 전체(최대 100만 행)를 훑다가 멈추지 않도록.
        stored = getattr(ws, "_cells", None)
        if stored is None:
            stored = {(c.row, c.column): c for row in ws.iter_rows() for c in row}
        values = {}
        for key, cell in list(stored.items()):
            text = _cell_text(cell.value, getattr(cell, "number_format", None), date1904)
            if text:
                link = getattr(cell, "hyperlink", None)
                if link is not None and (link.target or link.location):
                    text = _link_text(text, link.target, link.location)
                values[key] = text
        merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
        hidden_rows = {r for r, dim in ws.row_dimensions.items() if dim.hidden}
        hidden_cols = set()
        for letter, dim in ws.column_dimensions.items():
            if dim.hidden:
                try:
                    lo = dim.min or column_index_from_string(letter)
                    hi = dim.max or lo
                except Exception:
                    continue
                hidden_cols.update(range(lo, hi + 1))
        values, merges, note, _ = _drop_hidden(values, merges, hidden_rows, hidden_cols)
        sheets.append((ws.title, values, merges, note))
    return sheets, hidden_sheets


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
                    style_codes[str(i)] = custom.get(fid) or ("hh:mm" if fid in ("18", "19", "20", "21") else "yyyy-mm-dd")
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

    sheets, hidden_sheets = [], []
    for el in root.iter():
        if local(el.tag) != "sheet":
            continue
        if attr(el, "state") in ("hidden", "veryHidden"):
            hidden_sheets.append(attr(el, "name") or "")
            continue
        rid = attr(el, "id")
        if rid not in book_rels:
            continue
        part = book_rels[rid][1]
        if part not in names:
            continue
        values, merges = {}, []
        hidden_rows, hidden_cols = set(), set()
        row, col = 0, 0
        for event, node in ET.iterparse(z.open(part), events=("start", "end")):
            tag = local(node.tag)
            if event == "start":
                if tag == "row":
                    row = int(attr(node, "r") or row + 1)
                    col = 0
                    if attr(node, "hidden") in ("1", "true"):
                        hidden_rows.add(row)
                elif tag == "col" and attr(node, "hidden") in ("1", "true"):
                    try:
                        hidden_cols.update(range(int(attr(node, "min")), int(attr(node, "max")) + 1))
                    except (TypeError, ValueError):
                        pass
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
                                try:
                                    text = _date_text(number, style_codes.get(style), date1904)
                                except Exception:
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
        values, merges, note, _ = _drop_hidden(values, merges, hidden_rows, hidden_cols)
        sheets.append((attr(el, "name") or "", values, merges, note))
    return sheets, hidden_sheets


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
            texts = []  # 도형·글상자·단추 속 글
            for kind, drawing in rels_of(book_rels[attr(el, "id")][1]).values():
                if not kind.endswith("/drawing") or drawing not in names:
                    continue
                charts += sum(1 for t, _ in rels_of(drawing).values() if t.endswith("/chart"))
                tree = ET.fromstring(z.read(drawing))
                pictures += sum(1 for n in tree.iter() if local(n.tag) == "pic")
                for shape in tree.iter():
                    if local(shape.tag) != "sp":
                        continue
                    paras = []
                    for p in shape.iter():
                        if local(p.tag) == "p":
                            line = "".join(t.text or "" for t in p.iter() if local(t.tag) == "t").strip()
                            if line:
                                paras.append(re.sub(r"\s+", " ", line))
                    if paras:
                        texts.append(" ".join(paras))
            found[attr(el, "name") or ""] = (charts, pictures, texts)
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
        sheets, hidden_sheets = _xlsx_openpyxl(path)
    except Exception as first:
        try:
            sheets, hidden_sheets = _xlsx_raw(path)
        except Exception:
            raise first
    try:
        drawings = _xlsx_drawings(path)
    except Exception:
        drawings = {}
    out = []
    for title, values, merges, note in sheets:
        out.extend(_sheet_md(title, values, merges))
        if note:
            out.append(note)
        charts, pictures, texts = drawings.get(title, (0, 0, []))
        notes = [f"차트 {charts}개" if charts else "", f"그림 {pictures}개" if pictures else ""]
        if any(notes):
            out.append("[" + ", ".join(n for n in notes if n) + "]")
        if texts:
            out.append("도형·단추 속 글:\n\n" + "\n".join(f"- {t}" for t in texts))
    if _hidden_sheets_note(hidden_sheets):
        out.append(_hidden_sheets_note(hidden_sheets))
    from . import vba

    macros = vba.markdown(path)
    if macros:
        out.append(macros)
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
    out, hidden_sheets = [], []
    for index, sheet in enumerate(book.sheets()):
        if getattr(sheet, "visibility", 0):
            hidden_sheets.append(sheet.name)  # 숨긴 시트
            continue
        links = {}
        for link in getattr(sheet, "hyperlink_list", None) or []:
            for r in range(link.frowx, link.lrowx + 1):
                for c in range(link.fcolx, link.lcolx + 1):
                    links[(r, c)] = link
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
                    if code and _is_date_code(code):
                        try:
                            text = _date_text(value, code, book.datemode == 1)
                        except Exception:
                            text = ""
                    else:
                        text = ""
                    if not text:
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
                    link = links.get((r, c))
                    if link is not None and (link.url_or_path or link.textmark):
                        url = link.url_or_path if link.type in ("url", "unc") or (link.url_or_path or "").startswith(("http", "mailto")) else None
                        text = _link_text(text, url, None if url else (link.textmark or link.url_or_path))
                    values[(r + 1, c + 1)] = text
        merges = [(r0 + 1, c0 + 1, r1, c1) for r0, r1, c0, c1 in (sheet.merged_cells if formatted else [])]
        hidden_rows = {r + 1 for r, info in (getattr(sheet, "rowinfo_map", None) or {}).items() if getattr(info, "hidden", 0)}
        hidden_cols = {c + 1 for c, info in (getattr(sheet, "colinfo_map", None) or {}).items() if getattr(info, "hidden", 0)}
        values, merges, note, _ = _drop_hidden(values, merges, hidden_rows, hidden_cols)
        out.extend(_sheet_md(sheet.name, values, merges))
        if note:
            out.append(note)
    if _hidden_sheets_note(hidden_sheets):
        out.append(_hidden_sheets_note(hidden_sheets))
    from . import vba

    macros = vba.markdown(path)
    if macros:
        out.append(macros)
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
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):  # 엑셀 '유니코드 텍스트' 저장 등
        try:
            return data.decode("utf-16")
        except UnicodeDecodeError:
            pass
    for enc in ("utf-8-sig", "cp949", "utf-16"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _grid_md(rows, strip=False):
    if strip:
        rows = [[c.strip() for c in r] for r in rows]
    width = max((len(r) for r in rows), default=0)
    grid = [r + [""] * (width - len(r)) for r in rows]
    return table.render(table.from_grid(grid))


def convert_csv(path):
    text = read_text(path)
    tsv = str(path).lower().endswith(".tsv")
    rows = list(csv.reader(io.StringIO(text), delimiter="\t" if tsv else ","))
    # 엑셀이 탭 구분으로 저장하면 숫자 뒤에 빈칸이 붙는다 ("31,104,000 ")
    return _grid_md(rows, strip=tsv)


def _tab_rows(text):
    """탭으로 나눈 표처럼 보이면 행 목록, 아니면 None (엑셀의 '텍스트(탭으로 분리)' 저장 파일 등)."""
    if "\t" not in text:
        return None
    try:
        rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
    except csv.Error:
        return None
    rows = [r for r in rows if any(c.strip() for c in r)]
    if len(rows) < 2:
        return None
    multi = [r for r in rows if len(r) > 1]
    if len(multi) < 0.8 * len(rows):
        return None
    counts = {}
    for r in multi:
        counts[len(r)] = counts.get(len(r), 0) + 1
    if max(counts.values()) < 0.6 * len(multi):
        return None
    return rows


TXT_LIMIT = 50_000_000


def _indent(line):
    """줄 앞 들여쓰기를 보이는 빈칸으로 (마크다운이 코드 블록으로 읽지 않게)."""
    body = line.lstrip(" \t")
    lead = line[: len(line) - len(body)]
    width = sum(4 if ch == "\t" else 1 for ch in lead)
    return "\u00a0" * width + body


def convert_txt(path):
    """텍스트 파일. 탭으로 나뉜 표면 표로, 아니면 줄바꿈을 살린 글로."""
    if os.path.getsize(path) > TXT_LIMIT:
        raise ValueError(f"텍스트 파일이 너무 큽니다 ({TXT_LIMIT // 1_000_000}MB 초과). 기록(로그) 파일은 변환하지 않습니다.")
    text = read_text(path).replace("\r\n", "\n").replace("\r", "\n")
    rows = _tab_rows(text)
    if rows is not None:
        return _grid_md(rows, strip=True)
    paragraphs = []
    for block in re.split(r"\n[ \t]*\n", text):
        lines = [ln.rstrip() for ln in block.split("\n")]
        lines = [_indent(ln.replace("<", "&lt;")) for ln in lines if ln.strip()]
        if lines:
            paragraphs.append("  \n".join(lines))  # 원래 줄바꿈을 그대로 보이게
    return "\n\n".join(paragraphs)


# ---------------------------------------------------------------- 워드

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_TXBX = (W + "txbxContent",)


def _docx_iter(el, tag, skip=()):
    """el 아래의 tag 요소들. 호환용 사본(mc:Fallback — 같은 글상자가 옛 형식으로 한 번 더 들어 있음)과
    skip 에 든 요소의 안쪽은 보지 않는다."""
    for child in el:
        if child.tag == _MC_FALLBACK or child.tag in skip:
            continue
        if child.tag == tag:
            yield child
        yield from _docx_iter(child, tag, skip)


def _docx_tiny(box):
    """눈에 보이지 않을 만큼 작은 글상자(가로·세로 3mm 미만). PDF를 워드로 바꾼 파일에 흔한 찌꺼기."""
    node = box.getparent() if hasattr(box, "getparent") else None
    while node is not None:
        if node.tag.endswith("}wsp"):
            for ext in node.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
                try:
                    return int(ext.get("cx", "0")) < 100000 and int(ext.get("cy", "0")) < 100000
                except ValueError:
                    return False
            return False
        node = node.getparent()
    return False


def _docx_boxes(p):
    """문단에 붙은 글상자 속 문단·표들 (글상자 안의 글상자도 차례로)."""
    out = []
    for box in _docx_iter(p, W + "txbxContent"):
        if _docx_tiny(box):
            continue
        for inner in box:
            if inner.tag == W + "tbl":
                out.append(inner)
            elif inner.tag == W + "p":
                out.append(inner)
                out.extend(_docx_boxes(inner))
    return out


def _docx_runs(p):
    """문단의 글자를 굵게/기울임 표시와 함께 이어 붙인다 (글상자 속 글은 따로 다룬다)."""
    pieces = []  # (text, bold, italic)
    for r in _docx_iter(p, W + "r", _TXBX):
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
                    t for t in (_docx_text(p, markup=False) for p in _docx_iter(tc, W + "p")) if t
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
        _docx_para(el, styles, out)
        for inner in _docx_boxes(el):
            if inner.tag == W + "tbl":
                rendered = _docx_table(inner)
                if rendered:
                    out.append(rendered)
            else:
                _docx_para(inner, styles, out)
    return "\n\n".join(out)


def _docx_para(el, styles, out):
    """본문 문단 하나를 마크다운으로 바꿔 out 에 붙인다."""
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
        return
    text = _docx_text(el)
    if not text:
        return
    if is_list or low.startswith("list"):
        marker = "1. " if "number" in low else "- "
        item = marker + text.replace("\n", " ")
        if out and re.match(r"(- |1\. )", out[-1].splitlines()[-1]):
            out[-1] += "\n" + item
        else:
            out.append(item)
    else:
        out.append(text.replace("\n", "  \n"))


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


_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_ROMAN = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
          (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"))


def _roman(n):
    out = ""
    for value, letters in _ROMAN:
        while n >= value:
            out += letters
            n -= value
    return out


def _alpha(n):
    out = ""
    while n > 0:
        n, rest = divmod(n - 1, 26)
        out = chr(97 + rest) + out
    return out


def _num_label(kind, n):
    """자동 번호 종류(buAutoNum type)에 맞는 번호 글자."""
    kind = kind or "arabicPeriod"
    if kind.startswith("alphaLc"):
        core = _alpha(n)
    elif kind.startswith("alphaUc"):
        core = _alpha(n).upper()
    elif kind.startswith("romanLc"):
        core = _roman(n)
    elif kind.startswith("romanUc"):
        core = _roman(n).upper()
    elif kind.startswith("circleNum") and 1 <= n <= 20:
        return chr(0x2460 + n - 1)
    else:
        core = str(n)
    if kind.endswith("ParenBoth"):
        return f"({core})"
    if kind.endswith("ParenR"):
        return f"{core})"
    if kind.endswith("Plain"):
        return core
    return f"{core}."


def _slide_target(part, rid):
    """글자·도형에 건 링크가 가리키는 곳: ('url', 주소) 또는 ('slide', 번호)."""
    try:
        rel = part.rels[rid]
    except Exception:
        return None
    if rel.is_external:
        return ("url", rel.target_ref)
    try:
        target = rel.target_part
        slides = part.package.presentation_part.presentation.slides
        for i, sl in enumerate(slides, 1):
            if sl.part is target:
                return ("slide", i)
    except Exception:
        pass
    return None


def _linked(text, target):
    if not target or not text.strip():
        return text
    kind, where = target
    if kind == "url":
        return f"[{text}]({where.replace(' ', '%20').replace(')', '%29')})"
    return f"{text} (→ 슬라이드 {where})"


def _para_md(para, part):
    """문단 하나의 글: 링크는 [글](주소), 한 문단 안에서 일부만 굵게·기울임이면 **·* 로 표시한다."""
    pieces = []  # (글, 굵게, 기울임, 링크)
    for node in para._p:
        tag = node.tag
        if tag == _A + "r" or tag == _A + "fld":
            t = node.find(_A + "t")
            text = (t.text or "") if t is not None else ""
            rpr = node.find(_A + "rPr")
            bold = italic = False
            link = None
            if rpr is not None:
                bold = rpr.get("b") in ("1", "true")
                italic = rpr.get("i") in ("1", "true")
                click = rpr.find(_A + "hlinkClick")
                if click is not None and click.get(_R_ID):
                    link = _slide_target(part, click.get(_R_ID))
            pieces.append([text.replace("\x0b", " "), bold, italic, link])
        elif tag == _A + "br":
            pieces.append([" ", False, False, None])
    if not pieces:
        return _para_text(para)
    solid = [p for p in pieces if p[0].strip()]
    mixed_b = len({p[1] for p in solid}) > 1
    mixed_i = len({p[2] for p in solid}) > 1
    # 이웃한 같은 모양 조각을 합친다 (한 낱말이 여러 조각으로 저장되는 일이 흔하다)
    merged = []
    for p in pieces:
        key = (p[1] and mixed_b, p[2] and mixed_i, p[3])
        # 빈칸만 있는 조각은 앞 조각의 모양을 따른다 (굵은 두 낱말 사이 빈칸 때문에 ** 가 쪼개지지 않게)
        if merged and (merged[-1][1] == key or (not p[0].strip() and p[3] is None)):
            merged[-1][0] += p[0]
        else:
            merged.append([p[0], key])
    out = ""
    for i, (text, (bold, italic, link)) in enumerate(merged):
        core = text.strip()
        if not core:
            out += text
            continue
        lead = text[: len(text) - len(text.lstrip())]
        tail = text[len(text.rstrip()):]
        mark = ("**" if bold else "") + ("*" if italic else "")
        before = out[-1:] if not lead else " "
        after = (merged[i + 1][0][:1] if i + 1 < len(merged) else "") if not tail else " "
        # 마크다운은 문장부호로 끝나는 강조 바로 뒤에 글자가 붙으면 강조로 읽지 않는다: 그럴 땐 표시하지 않는다
        if mark and ((not core[-1].isalnum() and after.isalnum()) or (not core[0].isalnum() and before.isalnum())):
            mark = ""
        core = f"{mark}{core}{mark[::-1]}"
        out += lead + _linked(core, link) + tail
    return re.sub(r" {2,}", " ", out).strip()


def _bullet_kind(para):
    """문단에 직접 지정한 글머리 모양: ('none'|'char'|'num', 번호 종류, 시작 번호) 또는 None(물려받음)."""
    ppr = para._p.find(_A + "pPr")
    if ppr is None:
        return None
    if ppr.find(_A + "buNone") is not None:
        return ("none", None, None)
    num = ppr.find(_A + "buAutoNum")
    if num is not None:
        start = num.get("startAt")
        return ("num", num.get("type"), int(start) if start and start.isdigit() else None)
    if ppr.find(_A + "buChar") is not None or ppr.find(_A + "buBlip") is not None:
        return ("char", None, None)
    return None


def _shape_link(shape):
    """도형 전체에 건 링크(누르면 다른 슬라이드·주소로 이동)."""
    try:
        cnv = shape._element.find(".//" + "{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
        click = cnv.find(_A + "hlinkClick") if cnv is not None else None
        if click is not None and click.get(_R_ID):
            return _slide_target(shape.part, click.get(_R_ID))
    except Exception:
        pass
    return None


def _text_frame_md(shape):
    """글상자 하나를 마크다운으로. 목록이면 글머리표·번호, 아니면 줄글로 만든다."""
    try:
        part = shape.part
    except Exception:
        part = None
    paras = []
    for p in shape.text_frame.paragraphs:
        text = _para_md(p, part) if part is not None else _para_text(p)
        if text:
            paras.append((p.level or 0, text, _bullet_kind(p)))
    if not paras:
        return ""
    link = _shape_link(shape)
    if link is not None and not any("](" in t or "(→ 슬라이드" in t for _, t, _ in paras):
        last = paras[-1]
        paras[-1] = (last[0], _linked(last[1], link), last[2])
    glyphs = sum(1 for _, t, _ in paras if _BULLET.match(t))
    # 글머리 기호를 글자로 직접 쓴 목록 (한 문단짜리 '▶ 이동' 단추 글은 목록이 아니다)
    marked = glyphs >= 1 and len(paras) >= 2
    if not marked and any(k is not None and k[0] in ("char", "num") for _, _, k in paras):
        return _explicit_list_md(paras)
    body = _ph_type(shape) in ("BODY", "OBJECT", "SUBTITLE") and len(paras) > 1
    if not body and not marked:
        # 목록이 아닌 글상자: 문단마다 줄을 바꿔 한 덩어리로
        return "  \n".join(t.replace("\n", "  \n") for _, t, _ in paras)
    lines, seen_marker = [], False
    for level, text, kind in paras:
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


def _explicit_list_md(paras):
    """글머리표·자동 번호를 직접 지정한 문단들. 번호는 단계별로 이어 세고, 시작 번호가 있으면 따른다."""
    lines = []
    counters = {}  # 단계 -> (번호 종류, 지금 번호)
    in_list = False
    for level, text, kind in paras:
        flat = " ".join(text.split("\n"))
        for deeper in [lv for lv in counters if lv > level]:
            del counters[deeper]
        if kind is None or kind[0] == "none":
            if in_list:
                lines.append("")
            lines.append(flat)
            lines.append("")
            in_list = False
            continue
        indent = "   " * level
        if kind[0] == "char":
            lines.append(indent + "- " + flat)
        else:
            _, style, start = kind
            prev = counters.get(level)
            if start is not None:
                n = start
            elif prev is not None and prev[0] == style:
                n = prev[1] + 1
            else:
                n = 1
            counters[level] = (style, n)
            lines.append(f"{indent}{_num_label(style, n)} {flat}")
        in_list = True
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


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
