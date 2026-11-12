"""PDF 변환 (텍스트가 들어 있는 PDF 전용, OCR 없음).

쪽 범위 단위로 나눠 호출할 수 있어 여러 프로세스로 동시에 처리할 수 있다.
"""
import re

from . import table
from .table import Cell

SNAP_X = 6.0   # 이보다 가까운 세로선은 같은 열 경계로 본다 (셀 안쪽 여백 칸 제거)
SNAP_Y = 2.0
_NUMBER = re.compile(r"^[△▽\-+(]?\d{1,3}(,\d{3})+(\.\d+)?[%)]?$")
_TWO_NUMBERS = re.compile(r"^(△?\d{1,3}(?:,\d{3})+)(△?\d{1,3}(?:,\d{3})+)$")
_CID = re.compile(r"\(cid:\d+\)")  # 글꼴에 문자 정보가 없어 읽을 수 없는 글자

EMPTY_NOTE = (
    "> 이 PDF에서는 글자를 찾지 못했습니다. 스캔한 이미지로 된 PDF로 보이며, "
    "글자 인식(OCR)은 지원하지 않습니다."
)


def page_count(path):
    import pdfplumber

    with pdfplumber.open(path) as pdf:
        return len(pdf.pages)


# ---------------------------------------------------------------- 쪽 다듬기

def _mapped(page, objects, bbox=None):
    """객체 목록을 바꿔 끼운 쪽 객체를 만든다."""
    from pdfplumber.page import DerivedPage

    class Mapped(DerivedPage):
        def __init__(self, parent):
            self.bbox = tuple(bbox or parent.bbox)
            super().__init__(parent)

        @property
        def objects(self):
            return objects

    return Mapped(page)


def _clean_page(page):
    """쪽 밖에 놓인 도형을 버리고, 그림자·윤곽 효과로 여러 번 겹쳐 찍힌 글자를 하나로 줄인다."""
    x0, top, x1, bottom = page.bbox
    objects, changed = {}, False
    for kind, items in page.objects.items():
        kept = [
            o for o in items
            if not ("x0" in o and (o["x1"] < x0 or o["x0"] > x1 or o["bottom"] < top or o["top"] > bottom))
        ]
        if kind == "char":
            seen, unique = {}, []
            for ch in kept:
                key = (ch["text"], ch.get("fontname"), round(ch.get("size", 0), 1))
                near = False
                for px, pt in seen.get(key, ()):
                    if abs(ch["x0"] - px) < 0.6 * max(ch["width"], 1) and abs(ch["top"] - pt) < 0.6 * max(ch["height"], 1):
                        near = True
                        break
                if near and not ch["text"].isspace():
                    continue
                seen.setdefault(key, []).append((ch["x0"], ch["top"]))
                unique.append(ch)
            kept = unique
        changed = changed or len(kept) != len(items)
        objects[kind] = kept
    return _mapped(page, objects) if changed else page


# 글꼴에 문자 정보(ToUnicode)가 없으면 글자가 (cid:번호)로만 나온다.
# 숫자 글꼴은 대개 "문자 코드 - 일정한 값"이 글자 번호이므로, 그 값을 찾아 복원한다.
_CID_ONE = re.compile(r"\(cid:(\d+)\)")
_DECODED_OK = re.compile(r"^[-△(]?(\d{1,3}(,\d{3})+|\d{1,4})(\.\d+)?[%)]?$|^-$")
_FAMILY = re.compile(r"^[A-Z]{6}\+|[-,]?(Bold|Regular|Medium|Light|Semi|Demi|Extra|Black|Thin)+$", re.I)


def _family(fontname):
    name = fontname or ""
    while True:
        new = _FAMILY.sub("", name)
        if new == name:
            return name
        name = new


def _decode(cids, shift):
    return "".join(chr(c + shift) if 33 <= c + shift <= 126 else " " for c in cids)


def _find_shift(words):
    """낱말(글자 번호 목록)들이 가장 그럴듯한 숫자가 되는 이동값을 찾는다. 없으면 None."""
    if len(words) < 8:
        return None
    best, best_score = None, 0.0
    for shift in range(0, 96):
        ok = sum(1 for w in words if _DECODED_OK.match(_decode(w, shift).strip()))
        score = ok / len(words)
        if score > best_score:
            best, best_score = shift, score
    if best_score < 0.9:
        return None
    decoded = [_decode(w, best).strip() for w in words]
    # 쉼표가 든 금액이 실제로 있어야 숫자 글꼴로 인정한다 (우연히 맞아떨어진 경우 방지)
    if sum(1 for d in decoded if "," in d) < 3 and len(set("".join(decoded)) & set("0123456789")) < 6:
        return None
    return best


def _restore_glyphs(page, tol, shifts, ytol=3):
    """(cid:번호)로만 읽히는 숫자를 복원한다. (바뀐 쪽, 복원한 글자 수)를 돌려준다."""
    from pdfplumber.utils.text import extract_words

    chars = page.chars
    by_font = {}
    for ch in chars:
        if _CID_ONE.fullmatch(ch["text"]):
            by_font.setdefault(ch.get("fontname"), []).append(ch)
    if not by_font:
        return page, 0
    for font, items in by_font.items():
        fam = _family(font)
        if fam in shifts:
            continue
        try:
            words = [
                [int(n) for n in _CID_ONE.findall(w["text"])]
                for w in extract_words(items, x_tolerance=tol, y_tolerance=ytol)
            ]
        except Exception:
            words = []
        shift = _find_shift([w for w in words if w])
        if shift is not None:
            shifts[fam] = shift
    restored = 0
    new_chars = []
    for ch in chars:
        m = _CID_ONE.fullmatch(ch["text"])
        shift = shifts.get(_family(ch.get("fontname"))) if m else None
        if shift is not None:
            ch = dict(ch)
            ch["text"] = _decode([int(m.group(1))], shift)
            restored += 1
        new_chars.append(ch)
    if not restored:
        return page, 0
    objects = dict(page.objects)
    objects["char"] = new_chars
    return _mapped(page, objects), restored


# ---------------------------------------------------------------- 가로로 누운 쪽

def _rotated_page(page, clockwise):
    """글자가 90도 돌아가 있는 쪽을 똑바로 세운 쪽 객체로 만든다."""
    width, height = float(page.width), float(page.height)

    def point(x, top):
        if clockwise:
            return (height - top, x)
        return (top, width - x)

    def turn(obj):
        if not all(k in obj for k in ("x0", "x1", "top", "bottom")):
            return obj
        new = dict(obj)
        xa, ta = point(obj["x0"], obj["top"])
        xb, tb = point(obj["x1"], obj["bottom"])
        new["x0"], new["x1"] = min(xa, xb), max(xa, xb)
        new["top"], new["bottom"] = min(ta, tb), max(ta, tb)
        new["width"] = new["x1"] - new["x0"]
        new["height"] = new["bottom"] - new["top"]
        new["doctop"] = new["top"]
        new["y0"] = width - new["bottom"]
        new["y1"] = width - new["top"]
        if "pts" in obj:
            new["pts"] = [point(px, py) for px, py in obj["pts"]]
        if "upright" in obj:
            new["upright"] = not obj["upright"]
            m = obj.get("matrix")
            if m and not obj["upright"]:
                # 낱말을 묶을 때 글자 방향을 행렬에서 읽으므로, 바로 세운 글자에 맞게 고친다
                scale = abs(m[1]) or abs(m[2]) or 1
                new["matrix"] = (scale, 0, 0, scale, new["x0"], new["y0"])
        return new

    objects = {kind: [turn(o) for o in items] for kind, items in page.objects.items()}
    return _mapped(page, objects, (0, 0, height, width))


def _upright_page(page):
    chars = page.chars
    if len(chars) < 20:
        return page, False
    sideways = [c for c in chars if not c.get("upright", True)]
    if len(sideways) < len(chars) * 0.6:
        return page, False
    up = sum(1 for c in sideways if (c.get("matrix") or (0, 1))[1] > 0)
    rotated = _rotated_page(page, clockwise=up >= len(sideways) / 2)
    # 누운 글자에 딸린 공백 문자는 위치가 부정확해 숫자를 쪼개므로 버리고, 간격으로 단어를 나눈다
    return rotated.filter(
        lambda o: o.get("upright", True) and not (o.get("text") or "x").isspace()
    ), True


# ---------------------------------------------------------------- 표

def _snap(values, tolerance):
    ordered = sorted(set(values))
    index, mapping, last = -1, {}, None
    for v in ordered:
        if last is None or v - last > tolerance:
            index += 1
        mapping[v] = index
        last = v
    return mapping


def _cell_texts(tbl, chars, tol, ytol=3):
    """글자를 그 글자가 들어 있는 가장 작은 셀에 배정해 셀별 글을 만든다.

    세로선이 일부만 그어진 표에서는 큰 셀이 작은 셀들을 덮고 있어,
    단순히 셀 영역으로 자르면 아래 칸들의 글이 큰 셀에 함께 딸려 온다.
    """
    from bisect import bisect_right

    from pdfplumber.utils import extract_text
    from pdfplumber.utils.text import extract_words

    boxes = sorted(set(tbl.cells), key=lambda b: (b[2] - b[0]) * (b[3] - b[1]), reverse=True)
    xs = sorted({v for b in boxes for v in (b[0], b[2])})
    ys = sorted({v for b in boxes for v in (b[1], b[3])})
    xi = {v: i for i, v in enumerate(xs)}
    yi = {v: i for i, v in enumerate(ys)}
    slot = {}
    for b in boxes:  # 큰 셀부터 채우고 작은 셀이 덮어쓴다
        for i in range(xi[b[0]], xi[b[2]]):
            for j in range(yi[b[1]], yi[b[3]]):
                slot[(i, j)] = b
    # 선이 일부만 있어 칸으로 잡히지 않은 자리는 격자 한 칸씩을 임시 칸으로 삼는다
    made = set()
    for i in range(len(xs) - 1):
        for j in range(len(ys) - 1):
            if (i, j) not in slot:
                b = (xs[i], ys[j], xs[i + 1], ys[j + 1])
                slot[(i, j)] = b
                made.add(b)

    def cell_at(obj):
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return slot.get((bisect_right(xs, cx) - 1, bisect_right(ys, cy) - 1))

    x0, top, x1, bottom = tbl.bbox
    inside = [
        ch for ch in chars
        if x0 <= (ch["x0"] + ch["x1"]) / 2 <= x1 and top <= (ch["top"] + ch["bottom"]) / 2 <= bottom
    ]
    # 금액이 칸 경계선에 걸쳐 인쇄된 경우 숫자가 두 칸으로 쪼개지지 않도록,
    # 온전한 숫자 낱말은 낱말 중심이 놓인 칸에 통째로 넣는다
    owner = {}
    try:
        for gap in sorted({tol, max(tol, 3)}):
            for word in extract_words(inside, x_tolerance=gap, y_tolerance=ytol, return_chars=True):
                if any(id(ch) in owner for ch in word["chars"]):
                    continue
                pieces = [word["chars"]]
                glued = _TWO_NUMBERS.match(word["text"])
                if glued and len(word["chars"]) == len(word["text"]):
                    # 좁은 칸에서 넘친 두 금액이 붙어 한 낱말이 된 경우 둘로 가른다
                    cut = len(glued.group(1))
                    pieces = [word["chars"][:cut], word["chars"][cut:]]
                for piece in pieces:
                    box = {
                        "x0": min(c["x0"] for c in piece), "x1": max(c["x1"] for c in piece),
                        "top": min(c["top"] for c in piece), "bottom": max(c["bottom"] for c in piece),
                    }
                    b = cell_at(box)
                    if b is None:
                        continue
                    text = "".join(c["text"] for c in piece)
                    # 온전한 숫자이거나, 경계선 너머로 한두 글자만 삐져나온 낱말이면 통째로 넣는다
                    stray = sum(1 for ch in piece if cell_at(ch) is not b)
                    if stray and (_NUMBER.match(text) or stray <= 2):
                        for ch in piece:
                            owner[id(ch)] = b
    except Exception:
        owner = {}
    groups, orphans = {}, []
    for ch in inside:
        b = owner.get(id(ch)) or cell_at(ch)
        if b is not None:
            groups.setdefault(b, []).append(ch)
        elif not ch["text"].isspace():
            orphans.append(ch)
    def cell_text(cs):
        # 글자가 작은 칸은 줄 간격도 좁으므로, 줄을 가르는 허용치를 글자 크기에 맞춘다
        heights = sorted(c["bottom"] - c["top"] for c in cs)
        y_tol = min(3, max(0.8, heights[len(heights) // 2] * 0.4))
        return extract_text(cs, x_tolerance=tol, y_tolerance=y_tol).strip()

    texts = {b: cell_text(cs) for b, cs in groups.items()}
    width = max(x1 - x0, 1)
    # 임시 칸 하나에 여러 줄이 넓게 들어가면 세로선이 없는 표이므로 칸 구조를 믿을 수 없다
    sprawling = any(
        b in made and texts[b].count("\n") >= 2 and (b[2] - b[0]) >= width * 0.5 for b in texts
    )
    # 낱말이 표 영역의 좌우 경계에 걸쳐 있으면 표 영역이 실제보다 좁게 잡힌 것이다
    if not sprawling:
        mid = lambda ch: (ch["x0"] + ch["x1"]) / 2
        band = [ch for ch in chars if top <= (ch["top"] + ch["bottom"]) / 2 <= bottom]
        if len(band) > len(inside) and len(xs) <= 2:
            sprawling = True  # 한 칸짜리 표 옆에 다른 글이 있으면 표의 일부만 잡힌 것이다
        elif len(band) > len(inside):
            try:
                for word in extract_words(band, x_tolerance=tol, y_tolerance=ytol, return_chars=True):
                    sides = {x0 <= mid(ch) <= x1 for ch in word["chars"]}
                    if len(sides) == 2:
                        sprawling = True
                        break
            except Exception:
                pass
    total = sum(1 for ch in inside if not ch["text"].isspace())
    return texts, orphans, total, sprawling


def _merge_wrapped_rows(cells):
    """여러 줄로 나뉜 항목 이름 칸을 한 행으로 합친다.

    같은 행 범위를 통째로 차지하는 칸(금액 등)이 그 범위 안에서 줄마다 나뉜 칸보다
    많으면, 나뉜 칸들은 한 항목이 줄바꿈된 것으로 본다.
    """
    spans = sorted({(c.row, c.row + c.rowspan - 1) for c in cells if c.rowspan > 1 and c.text},
                   key=lambda s: s[1] - s[0])
    for a, b in spans:
        whole = [c for c in cells if c.text and c.row == a and c.row + c.rowspan - 1 == b]
        parts = [c for c in cells if c.text and a <= c.row and c.row + c.rowspan - 1 <= b
                 and not (c.row == a and c.row + c.rowspan - 1 == b)]
        cols = {c.col for c in parts}
        if not parts or len(whole) <= len(cols):
            continue
        for col in cols:
            group = sorted((c for c in parts if c.col == col), key=lambda c: c.row)
            first = group[0]
            first.text = "\n".join(c.text for c in group)
            first.rowspan = b - a + 1
            first.row = a
            for c in group[1:]:
                cells.remove(c)
    return cells


class _Broken(Exception):
    """선으로 찾은 칸에 들어가지 않는 글자가 많아 표 구조를 믿을 수 없다."""


def _table_cells(tbl, tol=3, lost=None, chars=None, ytol=3):
    raw = []
    if chars is not None:
        texts, orphans, total, sprawling = _cell_texts(tbl, chars, tol, ytol)
        if sprawling or len(orphans) > max(3, total * 0.03):
            raise _Broken()
        for bbox in set(tbl.cells) | set(texts):
            text, n = _CID.subn("", texts.get(bbox, ""))
            if lost is not None:
                lost[0] += n
            raw.append(([round(v, 1) for v in bbox], text))
    else:
        fallback = tbl.extract(x_tolerance=tol)
        for r, row in enumerate(tbl.rows):
            for j, bbox in enumerate(row.cells):
                if bbox is None:
                    continue
                text = fallback[r][j] if j < len(fallback[r]) else ""
                text, n = _CID.subn("", text or "")
                if lost is not None:
                    lost[0] += n
                raw.append(([round(v, 1) for v in bbox], text))
    if not raw:
        return []
    xmap = _snap([b[0] for b, _ in raw] + [b[2] for b, _ in raw], SNAP_X)
    ymap = _snap([b[1] for b, _ in raw] + [b[3] for b, _ in raw], SNAP_Y)
    cells = []
    for (x0, top, x1, bottom), text in raw:
        c0, c1, r0, r1 = xmap[x0], xmap[x1], ymap[top], ymap[bottom]
        if c1 <= c0 or r1 <= r0:
            if not text.strip():
                continue
            c1, r1 = max(c1, c0 + 1), max(r1, r0 + 1)
        cells.append(Cell(r0, c0, text, r1 - r0, c1 - c0))
    return _merge_wrapped_rows(cells)


# ---------------------------------------------------------------- 쪽 변환

def _region_fallback(page, bbox, tol, lost, ytol=3):
    """세로선이 없는 표: 글자 정렬로 칸을 추정하고, 그것도 글자를 놓치면 줄글로 내보낸다."""
    region = page.within_bbox(bbox, strict=False)
    lines = [
        _CID.subn("", ln["text"])
        for ln in region.extract_text_lines(x_tolerance=tol, y_tolerance=ytol)
    ]
    lost[0] += sum(n for _, n in lines)
    plain = "\n".join(t.strip() for t, _ in lines if t.strip())
    want = len(re.sub(r"\s", "", plain))
    try:
        grid = region.extract_table({
            "vertical_strategy": "text", "horizontal_strategy": "text",
            "text_x_tolerance": tol, "text_y_tolerance": ytol, "intersection_tolerance": 5,
        }) or []
        grid = [[_CID.sub("", c or "") for c in row] for row in grid]
        got = len(re.sub(r"\s", "", "".join("".join(r) for r in grid)))
        cols = max((sum(1 for c in r if c.strip()) for r in grid), default=0)
        amounts = lambda t: sorted(re.findall(r"\d{1,3}(?:,\d{3})+", t))
        whole = amounts(plain) == amounts(" | ".join(" | ".join(r) for r in grid))
        if grid and cols >= 2 and got >= want * 0.98 and whole:
            rendered = table.render_flat(table.from_grid(grid))
            if rendered:
                return rendered
    except Exception:
        pass
    return plain.replace("\n", "  \n")


def _drop_nested(tables):
    """다른 표의 영역 안에 완전히 들어 있는 표(칸 안 글상자가 표로 잡힌 것)를 뺀다."""
    def inside(a, b):
        return a[0] >= b[0] - 1 and a[1] >= b[1] - 1 and a[2] <= b[2] + 1 and a[3] <= b[3] + 1

    keep = []
    for t in tables:
        if not any(o is not t and inside(t.bbox, o.bbox) and not inside(o.bbox, t.bbox) for o in tables):
            keep.append(t)
    return keep


def _page_blocks(page, shifts=None):
    page, rotated = _upright_page(_clean_page(page))
    tol = 1.5 if rotated else 3
    lost = [0]
    heights = sorted(c["bottom"] - c["top"] for c in page.chars) or [10]
    # 글자가 아주 작은 쪽은 줄 간격도 좁으므로 줄을 가르는 허용치를 글자 크기에 맞춘다
    ytol = min(3, max(0.8, heights[len(heights) // 2] * 0.4))
    page, restored = _restore_glyphs(page, tol, shifts if shifts is not None else {}, ytol)
    try:
        tables = _drop_nested(page.find_tables({"text_x_tolerance": tol}))
    except Exception:
        tables = []
    chars = page.chars if tables else None
    prepared = []  # (영역, 변환된 글)
    for t in tables:
        bbox = t.bbox
        try:
            rendered = table.render_flat(_table_cells(t, tol, lost, chars, ytol))
        except _Broken:
            # 표 영역이 실제보다 좁게 잡히는 일이 많아 같은 높이의 쪽 전체 너비를 본다
            bbox = (page.bbox[0], t.bbox[1], page.bbox[2], t.bbox[3])
            rendered = _region_fallback(page, bbox, tol, lost, ytol)
        except Exception:
            grid = [[(c or "") for c in row] for row in t.extract()]
            rendered = table.render(table.from_grid(grid))
        prepared.append((bbox, rendered))
    boxes = [b for b, _ in prepared]

    def outside(obj):
        if "x0" not in obj or "top" not in obj:
            return True
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return not any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in boxes)

    text_page = page.filter(outside) if boxes else page
    items = []
    for line in text_page.extract_text_lines(x_tolerance=tol, y_tolerance=ytol):
        items.append((line["top"], "text", line))
    for bbox, rendered in prepared:
        items.append((bbox[1], "table", rendered))
    items.sort(key=lambda x: x[0])

    blocks, para, last_bottom = [], [], None
    for _, kind, obj in items:
        if kind == "table":
            if para:
                blocks.append("\n".join(para))
                para = []
            if obj:
                blocks.append(obj)
            last_bottom = None
            continue
        text, n = _CID.subn("", obj["text"])
        lost[0] += n
        text = text.strip()
        if not text:
            continue
        height = max(obj["bottom"] - obj["top"], 1)
        if para and last_bottom is not None and obj["top"] - last_bottom > height * 0.8:
            blocks.append("\n".join(para))
            para = []
        para.append(text)
        last_bottom = obj["bottom"]
    if para:
        blocks.append("\n".join(para))
    if lost[0]:
        blocks.insert(0, f"> 이 쪽에는 글꼴에 문자 정보가 없어 읽지 못한 글자가 {lost[0]}개 있습니다. 원본을 확인하세요.")
    if restored:
        blocks.insert(0, f"> 이 쪽의 숫자·기호 {restored}자는 글꼴에 문자 정보가 없어 글자 번호 규칙으로 복원했습니다. 중요한 값은 원본과 대조하세요.")
    return blocks


def convert_range(path, start, end):
    """start 이상 end 미만(0부터 세는 쪽 번호)을 변환해 [(쪽 번호, 글)]을 돌려준다."""
    import pdfplumber

    out = []
    shifts = {}  # 글꼴 계열별로 찾아낸 글자 번호 이동값 (같은 문서 안에서 재사용)
    with pdfplumber.open(path) as pdf:
        pages = pdf.pages
        for index in range(start, min(end, len(pages))):
            page = pages[index]
            try:
                blocks = _page_blocks(page, shifts)
            except Exception as exc:  # 한 쪽이 깨져도 나머지는 살린다
                blocks = [f"> ({index + 1}쪽을 변환하지 못했습니다: {exc})"]
            out.append((index + 1, "\n\n".join(blocks)))
            page.close()
    return out


def assemble(parts, total):
    """convert_range 결과들을 하나의 본문으로 합친다."""
    out = []
    for number, text in sorted(parts):
        if not text:
            continue
        if total > 1:
            out.append(f"<!-- {number}쪽 -->")
        out.append(text)
    if not out:
        return EMPTY_NOTE
    return "\n\n".join(out)


def convert_pdf(path):
    total = page_count(path)
    return assemble(convert_range(path, 0, total), total)
