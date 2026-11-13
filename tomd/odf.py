"""개방형 문서(ODF) 변환: 문서(.odt).

ODT는 zip 안의 content.xml에 본문이 들어 있는 표준 형식이다(한글·LibreOffice·워드에서 저장 가능).
"""
import re
import zipfile
from xml.etree import ElementTree as ET

from . import table
from .table import Cell

TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
XLINK = "http://www.w3.org/1999/xlink"


def _t(ns, name):
    return f"{{{ns}}}{name}"


def _styles(*roots):
    """글자 모양 스타일 이름 → (굵게, 기울임). 부모 스타일을 따라간다."""
    info, parent, hidden = {}, {}, set()
    numbered = {}  # 목록 스타일 이름 → 번호를 매기는 단계들
    for root in roots:
        if root is None:
            continue
        for ls in root.iter(_t(TEXT, "list-style")):
            levels = set()
            for lv in ls:
                if lv.tag == _t(TEXT, "list-level-style-number") and (lv.get(_t(STYLE, "num-format")) or "") != "":
                    levels.add(int(lv.get(_t(TEXT, "level"), "1") or 1))
            numbered[ls.get(_t(STYLE, "name"))] = levels
    for root in roots:
        if root is None:
            continue
        for st in root.iter(_t(STYLE, "style")):
            name = st.get(_t(STYLE, "name"))
            if not name:
                continue
            parent[name] = st.get(_t(STYLE, "parent-style-name"))
            props = st.find(_t(STYLE, "text-properties"))
            bold = italic = None
            if props is not None:
                # 한글은 '동아시아 글꼴' 굵기(font-weight-asian)에 따로 저장된다
                weights = [props.get(_t(FO, "font-weight")), props.get(_t(STYLE, "font-weight-asian")),
                           props.get(_t(STYLE, "font-weight-complex"))]
                weights = [w for w in weights if w]
                if weights:
                    bold = any(w == "bold" or (w.isdigit() and int(w) >= 600) for w in weights)
                slants = [props.get(_t(FO, "font-style")), props.get(_t(STYLE, "font-style-asian"))]
                slants = [x for x in slants if x]
                if slants:
                    italic = any(x in ("italic", "oblique") for x in slants)
                if props.get(_t(TEXT, "display")) == "none":
                    hidden.add(name)
            info[name] = (bold, italic)

    def resolve(name, depth=0):
        if not name or depth > 10:
            return False, False
        bold, italic = info.get(name, (None, None))
        if bold is None or italic is None:
            pb, pi = resolve(parent.get(name), depth + 1)
            bold = pb if bold is None else bold
            italic = pi if italic is None else italic
        return bool(bold), bool(italic)

    def is_hidden(name):
        return name in hidden

    return resolve, is_hidden, numbered


_TABLE_LINE = re.compile(r"\s*(\||</?(?:table|tr|td|th)\b)")


def _hard_breaks(text):
    """문단 안 줄바꿈을 마크다운 강제 줄바꿈(끝 빈칸 2개)으로. 글상자에서 온 표는 앞뒤를 빈 줄로 떼어
    표로 읽히게 하고, 표 줄에는 빈칸을 붙이지 않는다."""
    out = []
    lines = text.split("\n")
    for i, line in enumerate(lines):
        is_table = bool(_TABLE_LINE.match(line))
        if out:
            was_table = bool(_TABLE_LINE.match(out[-1]))
            if is_table != was_table:
                out.append("")
            elif not is_table:
                out[-1] += "  "
        out.append(line)
    return "\n".join(out)


class _Doc:
    def __init__(self, styles):
        self.resolve, self.is_hidden, self.numbered = styles
        self.notes = []
        self.images = 0

    # -- 글 조각
    def inline(self, el, markup=True):
        """문단 안의 글 (굵게·기울임·링크·각주 표시 포함). 같은 모양이 이어지는 조각은 합친다."""
        merged = []
        for text, bold, italic in self._segments(el, (False, False), markup):
            if merged and merged[-1][1:] == (bold, italic):
                merged[-1] = (merged[-1][0] + text, bold, italic)
            else:
                merged.append((text, bold, italic))
        out = []
        for text, bold, italic in merged:
            core = text.strip()
            if markup and core and "\n" not in core and (bold or italic):
                mark = ("**" if bold else "") + ("*" if italic else "")
                lead = text[: len(text) - len(text.lstrip())]
                tail = text[len(text.rstrip()):]
                out.append(f"{lead}{mark}{core}{mark[::-1]}{tail}")
            else:
                out.append(text)
        return "".join(out)

    def _segments(self, el, fmt, markup):
        """[(글, 굵게, 기울임)] 순서대로."""
        segs = [(el.text or "", *fmt)] if el.text else []
        for child in el:
            tag = child.tag
            if tag == _t(TEXT, "s"):
                segs.append((" " * int(child.get(_t(TEXT, "c"), "1") or 1), *fmt))
            elif tag == _t(TEXT, "tab"):
                segs.append((" ", *fmt))
            elif tag == _t(TEXT, "line-break"):
                segs.append(("\n", *fmt))
            elif tag == _t(TEXT, "span"):
                name = child.get(_t(TEXT, "style-name"))
                if not self.is_hidden(name):
                    bold, italic = self.resolve(name)
                    segs.extend(self._segments(child, (fmt[0] or bold, fmt[1] or italic), markup))
            elif tag == _t(TEXT, "a"):
                inner = self.inline(child, markup=False)
                href = child.get(_t(XLINK, "href")) or ""
                if markup and inner.strip() and re.match(r"(https?|mailto):", href, re.I):
                    segs.append((f"[{inner.strip()}]({href.replace(' ', '%20')})", False, False))
                else:
                    segs.append((inner, *fmt))
            elif tag == _t(TEXT, "note"):
                citation = child.find(_t(TEXT, "note-citation"))
                body = child.find(_t(TEXT, "note-body"))
                mark = (citation.text or "").strip() if citation is not None else ""
                mark = mark or str(len(self.notes) + 1)
                text = " ".join(self.blocks(body, markup=False)) if body is not None else ""
                self.notes.append((mark, text))
                segs.append((f"[^{mark}]" if markup else f"({mark})", False, False))
            elif tag in (_t(TEXT, "bookmark"), _t(TEXT, "bookmark-start"), _t(TEXT, "bookmark-end"),
                         _t(TEXT, "soft-page-break"), _t(OFFICE, "annotation"), _t(OFFICE, "annotation-end")):
                pass
            elif tag == _t(DRAW, "frame"):
                segs.append((self.frame(child), False, False))
            else:
                segs.extend(self._segments(child, fmt, markup))  # 필드(날짜·쪽 번호 등)는 보이는 글만
            if child.tail:
                segs.append((child.tail, *fmt))
        return segs

    def frame(self, frame):
        """그림·글상자. 글상자는 글을, 그림은 [그림]으로."""
        box = frame.find(_t(DRAW, "text-box"))
        if box is not None:
            return " ".join(self.blocks(box, markup=False))
        if frame.find(_t(DRAW, "image")) is not None:
            self.images += 1
            title = frame.find("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}title")
            name = (title.text or "").strip() if title is not None else ""
            return f"[그림: {name}]" if name else "[그림]"
        return ""

    # -- 블록
    def blocks(self, parent, markup=True, depth=0):
        out = []
        for el in parent:
            tag = el.tag
            if tag == _t(TEXT, "p"):
                text = _clean(self.inline(el, markup))
                if markup:
                    text = re.sub(r"<(?=[A-Za-z/!?])", "&lt;", text)  # 글 속 <b> 같은 글자가 태그로 읽히지 않게
                if text:
                    out.append(_hard_breaks(text) if markup else text)
            elif tag == _t(TEXT, "h"):
                text = _clean(self.inline(el, markup=False)).replace("\n", " ")
                if text:
                    level = int(el.get(_t(TEXT, "outline-level"), "1") or 1)
                    # 문서 제목(#)은 파일 이름이 쓰므로 한 단계 내려 쓴다
                    out.append(("#" * min(6, level + 1) + " " + text) if markup else text)
            elif tag == _t(TEXT, "list"):
                out.extend(self.list(el, markup, depth))
            elif tag == _t(TABLE, "table"):
                rendered = self.table(el)
                if rendered:
                    out.append(rendered if markup else re.sub(r"<[^>]+>", " ", rendered))
            elif tag in (_t(TEXT, "section"), _t(TEXT, "index-body"), _t(TEXT, "table-of-content"),
                         _t(TEXT, "alphabetical-index"), _t(TEXT, "illustration-index"), _t(TABLE, "table-header-rows")):
                out.extend(self.blocks(el, markup, depth))
            elif tag in (_t(TEXT, "index-title"),):
                out.extend(self.blocks(el, markup, depth))
            elif tag in (_t(TEXT, "sequence-decls"), _t(TEXT, "variable-decls"), _t(TEXT, "user-field-decls"),
                         _t(OFFICE, "forms"), _t(TEXT, "tracked-changes"), _t(TEXT, "table-of-content-source")):
                continue
            elif tag == _t(DRAW, "frame"):
                text = self.frame(el)
                if text:
                    out.append(text)
        return out

    def list(self, el, markup, depth, style=None):
        """목록. 목록 스타일이 번호 매기기면 1. 2. 로, 아니면 - 로. 들여쓰기로 단계를 나타낸다."""
        style = el.get(_t(TEXT, "style-name")) or style
        ordered = (depth + 1) in self.numbered.get(style, set())
        number = 0
        lines = []
        indent = "   " * depth
        for item in el:
            if item.tag not in (_t(TEXT, "list-item"), _t(TEXT, "list-header")):
                continue
            first = True
            number += 1
            start = item.get(_t(TEXT, "start-value"))
            if start and start.isdigit():
                number = int(start)
            marker = f"{number}. " if ordered else "- "
            for child in item:
                if child.tag == _t(TEXT, "list"):
                    lines.extend(self.list(child, markup, depth + 1, style))
                    continue
                for block in self.blocks([child], markup, depth + 1):
                    for line in block.split("\n"):
                        if first:
                            lines.append(f"{indent}{marker}{line}" if markup else line)
                            first = False
                        else:
                            lines.append(f"{indent}   {line}" if markup else line)
        return ["\n".join(lines)] if lines and markup and depth == 0 else lines

    def table(self, el):
        rows = []

        def collect(node):
            for child in node:
                if child.tag == _t(TABLE, "table-row"):
                    repeat = int(child.get(_t(TABLE, "number-rows-repeated"), "1") or 1)
                    row = []
                    for cell in child:
                        if cell.tag not in (_t(TABLE, "table-cell"), _t(TABLE, "covered-table-cell")):
                            continue
                        span_c = int(cell.get(_t(TABLE, "number-columns-spanned"), "1") or 1)
                        span_r = int(cell.get(_t(TABLE, "number-rows-spanned"), "1") or 1)
                        rep = int(cell.get(_t(TABLE, "number-columns-repeated"), "1") or 1)
                        covered = cell.tag == _t(TABLE, "covered-table-cell")
                        text = "" if covered else "\n".join(self.blocks(cell, markup=False))
                        for _ in range(min(rep, 1000)):
                            row.append((covered, text, span_r, span_c))
                    for _ in range(min(repeat, 1000) if any(t for _, t, _, _ in row) else 1):
                        rows.append(row)
                elif child.tag in (_t(TABLE, "table-header-rows"), _t(TABLE, "table-rows"), _t(TABLE, "table-row-group")):
                    collect(child)

        collect(el)
        cells = []
        for r, row in enumerate(rows):
            for c, (covered, text, rs, cs) in enumerate(row):
                if not covered:
                    cells.append(Cell(r, c, text, rs, cs))
        return table.render(cells, single="quote")


def _clean(text):
    return "\n".join(re.sub(r"[ \t　]{2,}", " ", line).strip() for line in text.split("\n")).strip()


def convert_odt(path):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        if "content.xml" not in names:
            raise ValueError("개방형 문서(ODT)의 본문을 찾지 못했습니다.")
        content = ET.fromstring(z.read("content.xml"))
        styles = ET.fromstring(z.read("styles.xml")) if "styles.xml" in names else None
    body = content.find(f"{_t(OFFICE, 'body')}/{_t(OFFICE, 'text')}")
    if body is None:
        raise ValueError("개방형 문서(ODT)의 본문을 찾지 못했습니다.")
    doc = _Doc(_styles(content, styles))
    blocks = doc.blocks(body)
    if doc.notes:
        blocks.append("\n".join(f"[^{mark}]: {text}" for mark, text in doc.notes))
    return "\n\n".join(b for b in blocks if b.strip())
