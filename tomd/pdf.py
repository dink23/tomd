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


def _near(a, b):
    return (
        abs(a["x0"] - b["x0"]) < 0.3 * max(a["width"], 1)
        and abs(a["top"] - b["top"]) < 0.3 * max(a["height"], 1)
    )


def _dedupe_chars(chars):
    """그림자·윤곽·굵게 효과로 같은 글자가 여러 번 겹쳐 찍힌 것을 한 번만 남긴다.

    사본은 직전 사본에서 아주 조금씩만 어긋나므로, 같은 글자가 거의 같은 자리(글자 크기의 30% 이내)에
    다시 나올 때만 사본으로 본다. 좁게 붙여 쓴 같은 글자(100.0%의 0 두 개)는 절반 넘게 떨어져
    있어 지워지지 않는다.
    """
    # 1) 서로 다른 글이 같은 자리에 겹쳐 찍힌 경우(쪽 번호가 두 개 찍힌 오류 등): 뒤의 것을 아래 줄로 내린다
    runs, current, prev = [], [], None
    for ch in chars:
        if not ch["text"].isspace():
            if prev is not None and (
                ch["x0"] < prev["x0"] - 0.5 * max(prev["width"], 1)
                or abs(ch["top"] - prev["top"]) > 0.5 * max(prev["height"], 1)
            ):
                runs.append(current)
                current = []
            prev = ch
        current.append(ch)
    runs.append(current)
    grid, shifted = {}, []
    cell = lambda c: (int(c["x0"] // 4), int(c["top"] // 4))
    for run in runs:
        solid = [c for c in run if not c["text"].isspace()]
        clash, bottom = False, 0.0
        if len(solid) >= 3:
            same = 0
            for c in solid:
                gx, gy = cell(c)
                for key in ((gx + i, gy + j) for i in (-1, 0, 1) for j in (-1, 0, 1)):
                    for other in grid.get(key, ()):
                        if _near(c, other):
                            if other["text"] == c["text"]:
                                same += 1
                            else:
                                clash = True
                                bottom = max(bottom, other["bottom"])
            # 대부분 같은 글자인데 일부만 다르면 다른 글이 겹친 것이다 (그림자 사본은 전부 같다)
            clash = clash and same >= 1
        if clash:
            shift = bottom - min(c["top"] for c in solid) + 1
            moved = []
            for c in run:
                c = dict(c)
                for k in ("top", "bottom", "doctop"):
                    if k in c:
                        c[k] += shift
                for k in ("y0", "y1"):
                    if k in c:
                        c[k] -= shift
                moved.append(c)
            run = moved
        for c in run:
            if not c["text"].isspace():
                grid.setdefault(cell(c), []).append(c)
        shifted.extend(run)

    # 2) 같은 글자가 거의 같은 자리에 다시 찍힌 사본을 지운다
    trails, out, prev = {}, [], None
    for ch in shifted:
        if ch["text"].isspace():
            out.append(ch)
            continue
        # 같은 줄에서 바로 앞 글자에 이어 일정하게 나아가며 찍힌 같은 글자(목차의 점선 등)는 사본이 아니다
        advancing = (
            prev is not None and prev["text"] == ch["text"]
            and abs(ch["top"] - prev["top"]) < 0.01 and ch["x0"] - prev["x0"] >= 0.75
        )
        prev = ch
        copy = False
        for trail in trails.setdefault(ch["text"], []):
            if _near(ch, trail):
                if advancing:
                    break
                trail["x0"], trail["top"] = ch["x0"], ch["top"]
                copy = True
                break
        if copy:
            continue
        trails[ch["text"]].append({"x0": ch["x0"], "top": ch["top"], "width": ch["width"], "height": ch["height"]})
        out.append(ch)
    return out


def _align_small_glyphs(chars):
    """가운뎃점·위첨자처럼 기준선이 다른 글자가 제 줄에서 떨어져 나가지 않도록 앞 글자의 줄에 맞춘다."""
    out, prev = [], None
    for ch in chars:
        if ch["text"].isspace():
            out.append(ch)
            continue
        if prev is not None and abs(ch["top"] - prev["top"]) > 0.5:
            overlap = min(ch["bottom"], prev["bottom"]) - max(ch["top"], prev["top"])
            adjacent = (
                ch["x0"] - prev["x1"] <= 0.6 * prev["height"]
                and ch["x0"] >= prev["x0"] - 0.5 * prev["height"]
            )
            similar = min(ch["height"], prev["height"]) >= 0.5 * max(ch["height"], prev["height"])
            if adjacent and similar and overlap >= 0.6 * min(ch["height"], prev["height"]):
                ch = dict(ch)
                shift = prev["top"] - ch["top"]
                ch["top"], ch["bottom"] = prev["top"], prev["bottom"]
                ch["height"] = ch["bottom"] - ch["top"]
                if "doctop" in ch:
                    ch["doctop"] += shift
                out.append(ch)
                # 줄 높이는 원래 기준 글자의 것을 유지하되, 이어지는 글자가 계속 맞춰지도록 오른쪽 끝만 옮긴다
                prev = dict(prev)
                prev["x1"] = max(prev["x1"], ch["x1"])
                prev["x0"] = ch["x0"]
                continue
        out.append(ch)
        prev = ch
    return out


_LETTER = re.compile(r"[가-힣A-Za-z]")


def _vertical_runs(chars):
    """한 글자씩 세로로 쌓아 쓴 글(책등, 세로 제목)을 찾아 글자 목록들로 돌려준다.

    낱말 사이의 빈칸 글자도 함께 담아 띄어쓰기를 살린다.
    """
    runs, current = [], []

    def close():
        solid = [c for c in current if not c["text"].isspace()]
        letters = sum(1 for c in solid if _LETTER.match(c["text"]))
        if len(solid) >= 5 and letters >= 0.8 * len(solid):
            while current and current[-1]["text"].isspace():
                current.pop()
            runs.append(list(current))
        current.clear()

    def same_column(a, b):
        return abs((b["x0"] + b["x1"]) / 2 - (a["x0"] + a["x1"]) / 2) < 0.3 * max(a["width"], 1)

    for ch in chars:
        solid = [c for c in current if not c["text"].isspace()]
        prev = solid[-1] if solid else None
        if ch["text"].isspace():
            if prev is not None and same_column(prev, ch) and 0 <= ch["top"] - prev["top"] <= 2.5 * prev["height"]:
                current.append(ch)
            continue
        if prev is not None:
            step = ch["top"] - prev["top"]
            spaced = current[-1]["text"].isspace()
            limit = (3.5 if spaced else 2.2) * prev["height"]
            if not (same_column(prev, ch) and 0.7 * prev["height"] <= step <= limit):
                close()
        current.append(ch)
    close()
    return runs


def _oblique(ch):
    """글자가 비스듬히 놓였는가 (원을 따라 휘어 쓴 글). 90도 단위로 돌린 글자는 아니다."""
    m = ch.get("matrix")
    if not m:
        return False
    a, b = abs(m[0]), abs(m[1])
    return min(a, b) > 0.08 * (a + b)


def _curved_runs(chars):
    """곡선을 따라 쓴 글자들을 글자 흐름 순서대로 묶는다.

    이런 글자는 높이가 제각각이라 가로 줄로 묶으면 한두 글자씩 흩어진다.
    """
    runs, run, bridge = [], [], []

    def close():
        if sum(1 for c in run if not c["text"].isspace()) >= 4:
            runs.append(list(run))
        run.clear()
        bridge.clear()

    for ch in chars:
        near = True
        if run:
            last = run[-1]
            size = max(last["height"], last["width"], ch["height"], 1)
            near = (abs(ch["x0"] - last["x0"]) < 2.5 * size and abs(ch["top"] - last["top"]) < 2.5 * size)
        if _oblique(ch):
            if not near:
                close()
            run.extend(bridge)
            bridge.clear()
            run.append(ch)
        elif run and near and len(bridge) < 3:
            bridge.append(ch)  # 곡선 꼭대기의 거의 바로 선 글자·빈칸은 이어 준다
        else:
            close()
    close()
    return runs


def _side_by_side(lines, tol, ytol, page_chars=()):
    """선 없이 좌우로 나란히 놓인 글 묶음(왼쪽 이름표 + 오른쪽 항목들)을 칸으로 가른다.

    모든 줄에 걸쳐 글자가 하나도 없는 세로 틈이 있으면 그 틈을 기준으로 나눈다.
    나눌 수 없으면 None, 나눴으면 칸별 글 목록을 돌려준다.
    """
    from pdfplumber.utils import extract_text

    every = [c for ln in lines for c in ln.get("chars", [])]
    chars = [c for c in every if not c["text"].isspace()]
    if chars and page_chars:
        # 줄에 딸린 글자 목록에는 빈칸 글자가 빠져 있으므로, 같은 영역의 빈칸 글자를 쪽에서 가져온다
        x0, x1 = min(c["x0"] for c in chars), max(c["x1"] for c in chars)
        top, bottom = min(c["top"] for c in chars), max(c["bottom"] for c in chars)
        every = every + [
            c for c in page_chars
            if c["text"].isspace() and x0 <= (c["x0"] + c["x1"]) / 2 <= x1
            and top <= (c["top"] + c["bottom"]) / 2 <= bottom
        ]
    if len(chars) < 8:
        return None
    size = sorted(c["bottom"] - c["top"] for c in chars)[len(chars) // 2]
    gap_min = max(15, 1.5 * size)
    spans = sorted((c["x0"], c["x1"]) for c in chars)
    cuts, reach = [], spans[0][1]
    for x0, x1 in spans[1:]:
        if x0 - reach >= gap_min:
            cuts.append((reach + x0) / 2)
        reach = max(reach, x1)
    if not cuts or len(cuts) > 3:
        return None
    columns = [[] for _ in range(len(cuts) + 1)]
    for c in every:  # 빈칸 글자도 함께 넣어야 낱말 사이 띄어쓰기가 살아난다
        columns[sum(1 for cut in cuts if (c["x0"] + c["x1"]) / 2 > cut)].append(c)
    texts = []
    for col in columns:
        text = extract_text(col, x_tolerance=tol, y_tolerance=ytol).strip()
        rows = [t.strip() for t in text.split("\n") if t.strip()]
        letters = len(re.sub(r"\s", "", text))
        # 번호·글머리표만 있는 칸(1. 2. / 가. 나.)은 목록의 내어쓰기이지 칸이 아니다
        if not rows or letters < 6 or letters / len(rows) < 3:
            return None
        # 오른쪽 끝에 따로 붙인 "(단위: 천원)" 같은 주석은 칸이 아니다
        if len(rows) == 1 and re.match(r"^[(\[〔<].*[)\]〕>]$", rows[0]):
            return None
        texts.append(rows)
    if sum(len(rows) for rows in texts) < 3:
        return None
    return texts


def _clean_page(page):
    """쪽 밖에 놓인 도형과 글을 버리고, 겹쳐 찍힌 글자와 줄에서 떨어진 작은 글자를 정리한다."""
    x0, top, x1, bottom = page.bbox
    objects, changed = {}, False
    for kind, items in page.objects.items():
        kept = [
            o for o in items
            if not ("x0" in o and (o["x1"] < x0 or o["x0"] > x1 or o["bottom"] < top or o["top"] > bottom))
        ]
        if kind == "char":
            cleaned = _dedupe_chars(_align_small_glyphs(kept))
            changed = changed or len(cleaned) != len(kept) or any(a is not b for a, b in zip(cleaned, kept))
            kept = cleaned
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
        return page, False, []
    sideways = [c for c in chars if not c.get("upright", True)]
    if len(sideways) < len(chars) * 0.6:
        return page, False, []
    # 누운 쪽에서 혼자 바로 서 있는 글(쪽 번호 등)은 돌리기 전에 따로 읽어 둔다
    try:
        extra = [
            ln["text"].strip()
            for ln in page.filter(lambda o: o.get("upright", True)).extract_text_lines()
            if ln["text"].strip()
        ]
    except Exception:
        extra = []
    up = sum(1 for c in sideways if (c.get("matrix") or (0, 1))[1] > 0)
    rotated = _rotated_page(page, clockwise=up >= len(sideways) / 2)
    # 누운 글자에 딸린 공백 문자는 위치가 부정확해 숫자를 쪼개므로 버리고, 간격으로 단어를 나눈다
    return rotated.filter(
        lambda o: o.get("upright", True) and not (o.get("text") or "x").isspace()
    ), True, extra


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


def _open_side_borders(page):
    """좌우 테두리 없이 가로선과 안쪽 세로선만 그은 표에 양옆 세로선을 보태 준다.

    이런 표는 칸의 네 귀퉁이가 닫히지 않아 선만으로는 표로 잡히지 않는다.
    """
    rows = {}
    for e in page.edges:
        if e.get("orientation") == "h" and e["x1"] - e["x0"] > 5:
            rows.setdefault(round(e["top"]), []).append((e["x0"], e["x1"]))
    spans = []  # (x0, x1, y) : 같은 높이에서 이어진 조각을 합친 가로선
    for y, parts in rows.items():
        parts.sort()
        a, b = parts[0]
        for x0, x1 in parts[1:]:
            if x0 <= b + 2:
                b = max(b, x1)
            else:
                spans.append((a, b, y))
                a, b = x0, x1
        spans.append((a, b, y))
    groups = []
    for a, b, y in sorted(spans):
        if b - a < 100:
            continue
        for g in groups:
            if abs(g[0] - a) <= 2 and abs(g[1] - b) <= 2:
                g[2].append(y)
                break
        else:
            groups.append([a, b, [y]])
    verticals = [e for e in page.edges if e.get("orientation") == "v" and e["bottom"] - e["top"] > 5]
    added = []
    for a, b, ys in groups:
        if len(ys) < 3:
            continue
        ys = sorted(set(ys))

        def spanned(x_ok, y1, y2):
            return sum(
                min(v["bottom"], y2) - max(v["top"], y1)
                for v in verticals if x_ok(v["x0"]) and v["bottom"] > y1 and v["top"] < y2
            ) >= 0.5 * (y2 - y1)

        for x in (a, b):
            run = []  # 테두리를 보탤 연속 구간
            for y1, y2 in list(zip(ys, ys[1:])) + [(None, None)]:
                # 칸을 나누는 안쪽 세로선이 지나가는 줄 사이에만 보탠다 (표와 표 사이는 잇지 않는다)
                need = (
                    y1 is not None
                    and spanned(lambda vx: a + 5 < vx < b - 5, y1, y2)
                    and not spanned(lambda vx: abs(vx - x) <= 3, y1, y2)
                )
                if need:
                    run.append((y1, y2))
                    continue
                if len(run) >= 2:
                    top, bottom = run[0][0], run[-1][1]
                    added.append({
                        "object_type": "line", "x0": x, "x1": x, "top": top, "bottom": bottom,
                        "width": 0, "height": bottom - top, "orientation": "v",
                    })
                run = []
    return added


def _is_white(color):
    if color is None:
        return True
    if isinstance(color, (int, float)):
        color = (color,)
    try:
        values = [float(v) for v in color]
    except (TypeError, ValueError):
        return True
    if len(values) == 4:
        return sum(values) < 0.05
    return all(v > 0.97 for v in values)


def _fills(page):
    """칸을 색으로 칠한 사각형들 [(x0, top, x1, bottom, 색)] — 그린 순서대로."""
    out = []
    for r in page.rects:
        if not r.get("fill") or _is_white(r.get("non_stroking_color")):
            continue
        if r["x1"] - r["x0"] < 3 or r["bottom"] - r["top"] < 3:
            continue
        color = r.get("non_stroking_color")
        if isinstance(color, (int, float)):
            color = (color,)
        out.append((r["x0"], r["top"], r["x1"], r["bottom"], tuple(round(float(v), 2) for v in color)))
    return out


def _mark_shaded(raw, fills, spans):
    """일정표처럼 빈칸을 색으로 칠해 뜻을 나타낸 표에서, 칠해진 빈칸에 ■를 넣는다.

    글이 든 칸에도 쓰인 색(머리글 바탕, 줄무늬)은 꾸밈이므로 표시하지 않는다.
    같은 행과 같은 열에 글이 있는 칸(행 이름·열 이름이 있는 자리)만 표시한다.
    spans는 칸마다의 (첫 행, 끝 행, 첫 열, 끝 열).
    """
    def color_at(b):
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
        found = None
        for x0, top, x1, bottom, color in fills:
            if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bottom + 1:
                found = color
        return found

    if not raw or not spans:
        return raw
    colors = [color_at(b) for b, _ in raw]
    first_row = min(sp[0] for sp in spans)
    rows, cols = set(), set()        # 한 줄·한 칸짜리 글 칸이 있는 행과 열 (행 이름·열 이름)
    row_colors, col_colors = {}, {}  # 글이 든 칸의 바탕색: 줄무늬·강조 열은 꾸밈으로 본다
    for (_, text), color, (r0, r1, c0, c1) in zip(raw, colors, spans):
        if not text.strip():
            continue
        if r1 - r0 == 1:
            rows.add(r0)
        if c1 - c0 == 1:
            cols.add(c0)
        if color:
            for r in range(r0, r1):
                row_colors.setdefault(r, set()).add(color)
            if r0 > first_row:  # 머리글 줄의 바탕색은 일정 막대와 같은 색일 수 있다
                for c in range(c0, c1):
                    col_colors.setdefault(c, set()).add(color)
    empty = [i for i, (_, text) in enumerate(raw) if not text.strip()]
    marks = []
    for i in empty:
        r0, r1, c0, c1 = spans[i]
        color = colors[i]
        if not color or r0 == first_row:
            continue
        if not rows.intersection(range(r0, r1)) or not cols.intersection(range(c0, c1)):
            continue
        if any(color in row_colors.get(r, ()) for r in range(r0, r1)):
            continue
        if any(color in col_colors.get(c, ()) for c in range(c0, c1)):
            continue
        marks.append(i)
    if len(marks) >= 3 and len(marks) < len(empty):
        for i in marks:
            raw[i] = (raw[i][0], "■")
    return raw


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
    groups, orphans, cell_of = {}, [], {}
    for ch in inside:
        b = owner.get(id(ch)) or cell_at(ch)
        if b is not None:
            groups.setdefault(b, []).append(ch)
            cell_of[id(ch)] = b
        elif not ch["text"].isspace():
            orphans.append(ch)
    # 표 테두리 밖으로 내어 쓴 글(행 이름, 제목 앞의 □ 등)은 같은 줄의 가장 가까운 칸에 넣는다
    claimed = set()
    inside_ids = {id(ch) for ch in inside}
    beside = [
        ch for ch in chars
        if id(ch) not in inside_ids and not ch["text"].isspace()
        and top <= (ch["top"] + ch["bottom"]) / 2 <= bottom
    ]
    if beside:
        anchors = [ch for ch in inside if id(ch) in cell_of and not ch["text"].isspace()]
        left = sorted((c for c in beside if c["x1"] <= x0 + c["width"]), key=lambda c: -c["x1"])
        right = sorted((c for c in beside if c["x0"] >= x1 - c["width"]), key=lambda c: c["x0"])
        for side, group in (("left", left), ("right", right)):
            for ch in group:
                cy = (ch["top"] + ch["bottom"]) / 2
                reach = 1.5 * max(ch["height"], 1)
                best = None
                for a in anchors:
                    if not (a["top"] <= cy <= a["bottom"]):
                        continue
                    gap = a["x0"] - ch["x1"] if side == "left" else ch["x0"] - a["x1"]
                    if -0.5 * ch["width"] <= gap <= reach and (best is None or gap < best[0]):
                        best = (gap, a)
                if best is not None:
                    b = cell_of[id(best[1])]
                    groups[b].append(ch)
                    cell_of[id(ch)] = b
                    claimed.add(id(ch))
                    anchors.append(ch)
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
        band = [
            ch for ch in chars
            if top <= (ch["top"] + ch["bottom"]) / 2 <= bottom and id(ch) not in claimed
        ]
        if len(band) > len(inside) and len(xs) <= 2:
            sprawling = True  # 한 칸짜리 표 옆에 다른 글이 있으면 표의 일부만 잡힌 것이다
        elif len(band) > len(inside):
            try:
                for word in extract_words(band, x_tolerance=tol, y_tolerance=ytol, return_chars=True):
                    sides = {x0 <= mid(ch) <= x1 or id(ch) in claimed for ch in word["chars"]}
                    if len(sides) == 2:
                        sprawling = True
                        break
            except Exception:
                pass
    total = sum(1 for ch in inside if not ch["text"].isspace())
    # 글자보다 좁은 칸에 글자가 들어가 있으면 표가 아니라 그라데이션 같은 장식 도형이다
    squeezed = sum(
        1 for ch in inside
        if not ch["text"].isspace() and id(ch) in cell_of
        and cell_of[id(ch)][2] - cell_of[id(ch)][0] < 0.9 * ch["width"]
    )
    if total and squeezed > 0.3 * total:
        sprawling = "decoration"
    return texts, orphans, total, sprawling, claimed


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


def _table_cells(tbl, tol=3, lost=None, chars=None, ytol=3, claimed=None, fills=None):
    raw = []
    if chars is not None:
        texts, orphans, total, sprawling, taken = _cell_texts(tbl, chars, tol, ytol)
        if sprawling or len(orphans) > max(3, total * 0.03):
            raise _Broken("decoration" if sprawling == "decoration" else "")
        if claimed is not None:
            claimed.update(taken)
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
    if fills:
        # 열·행 경계를 맞춘 뒤에도 남는 칸만 대상으로 한다 (칸 안쪽 여백 조각 제외)
        real = [i for i, (b, _) in enumerate(raw) if xmap[b[2]] > xmap[b[0]] and ymap[b[3]] > ymap[b[1]]]
        spans = [(ymap[raw[i][0][1]], ymap[raw[i][0][3]], xmap[raw[i][0][0]], xmap[raw[i][0][2]]) for i in real]
        marked = _mark_shaded([raw[i] for i in real], fills, spans)
        for i, item in zip(real, marked):
            raw[i] = item
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

def _region_fallback(page, bbox, tol, lost, ytol=3, plain_only=False):
    """세로선이 없는 표: 글자 정렬로 칸을 추정하고, 그것도 글자를 놓치면 줄글로 내보낸다."""
    def centered(obj):
        if "x0" not in obj or "top" not in obj:
            return False
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return bbox[0] <= cx <= bbox[2] and bbox[1] <= cy <= bbox[3]

    # 표 밖 글을 고를 때와 같은 기준(중심점)으로 골라야 경계에 걸친 글자가 빠지지 않는다
    region = page.filter(centered)
    lines = [
        _CID.subn("", ln["text"])
        for ln in region.extract_text_lines(x_tolerance=tol, y_tolerance=ytol)
    ]
    lost[0] += sum(n for _, n in lines)
    plain = "\n".join(t.strip() for t, _ in lines if t.strip())
    want = len(re.sub(r"\s", "", plain))
    if plain_only:
        return plain.replace("\n", "  \n")
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
        if grid and cols >= 2 and got >= want and whole:
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
    page = _clean_page(page)
    original = page
    # 세로로 쌓아 쓴 글은 가로 글과 섞이면 한 글자씩 흩어지므로 따로 떼어 한 줄로 만든다
    vertical = []
    try:
        runs = _vertical_runs(page.chars)
    except Exception:
        runs = []
    try:
        curved = _curved_runs(page.chars)
    except Exception:
        curved = []
    if curved:
        used = {id(c) for run in runs for c in run}
        curved = [run for run in curved if not any(id(c) in used for c in run)]
        for run in curved:
            text = re.sub(r"\s+", " ", "".join(c["text"] for c in run)).strip()
            vertical.append((min(c["top"] for c in run), text))
        gone = {id(c) for run in curved for c in run}
        page = page.filter(lambda o: id(o) not in gone)
    if runs:
        taken = {id(c) for run in runs for c in run}
        for run in runs:
            text, last = "", None
            for c in run:
                if c["text"].isspace():
                    if text and not text.endswith(" "):
                        text += " "
                    last = None
                    continue
                if last is not None and c["top"] - last["bottom"] > 0.5 * last["height"] and not text.endswith(" "):
                    text += " "
                text += c["text"]
                last = c
            vertical.append((run[0]["top"], text.strip()))
        page = page.filter(lambda o: id(o) not in taken)
    page, rotated, extra = _upright_page(page)
    if rotated:
        # 돌려 놓은 쪽에서는 떼어 낸 글의 자리를 맞출 수 없다. 버리면 글이 빠지므로
        # 떼어 내기 전의 쪽으로 되돌려 모든 글자를 본문 흐름에 남긴다.
        vertical = []
        if page is not original and (runs or curved):
            page, rotated, extra = _upright_page(original)
    tol = 1.5 if rotated else 3
    lost = [0]
    heights = sorted(c["bottom"] - c["top"] for c in page.chars) or [10]
    # 글자가 아주 작은 쪽은 줄 간격도 좁으므로 줄을 가르는 허용치를 글자 크기에 맞춘다
    ytol = min(3, max(0.8, heights[len(heights) // 2] * 0.4))
    page, restored = _restore_glyphs(page, tol, shifts if shifts is not None else {}, ytol)
    settings = {"text_x_tolerance": tol}
    try:
        tables = _drop_nested(page.find_tables(settings))
    except Exception:
        tables = []
    fills = []
    if not rotated:
        # 양옆 테두리가 없어 표로 잡히지 않은 구간에만 세로선을 보태 다시 찾는다.
        # 이미 표로 잡힌 구간은 건드리지 않는다 (잘 되던 표의 칸 구조가 바뀌지 않도록).
        try:
            def found(line):
                got = sum(
                    max(0, min(t.bbox[3], line["bottom"]) - max(t.bbox[1], line["top"]))
                    for t in tables if t.bbox[0] - 5 <= line["x0"] <= t.bbox[2] + 5
                )
                return got >= 0.6 * (line["bottom"] - line["top"])

            borders = [ln for ln in _open_side_borders(page) if not found(ln)]
            if borders:
                again = _drop_nested(page.find_tables(dict(settings, explicit_vertical_lines=borders)))
                # 보탠 선으로 새로 생긴 표가 온전할 때만 받아들인다 (장식 선이 많은 디자인 문서 보호)
                before = {tuple(round(v) for v in t.bbox) for t in tables}
                fresh = [t for t in again if tuple(round(v) for v in t.bbox) not in before]
                for t in fresh:
                    _table_cells(t, tol, [0], page.chars, ytol, set())
                if fresh:
                    tables = again
        except Exception:
            pass
        try:
            fills = _fills(page)
        except Exception:
            fills = []
    chars = page.chars if tables else None
    prepared = []  # (영역, 변환된 글)
    claimed = set()  # 표 밖에 있지만 표의 칸으로 가져간 글자
    for t in tables:
        bbox = t.bbox
        try:
            rendered = table.render_flat(_table_cells(t, tol, lost, chars, ytol, claimed, fills))
        except _Broken as broken:
            # 표 영역이 실제보다 좁게 잡히는 일이 많아 같은 높이의 쪽 전체 너비를 본다
            bbox = (page.bbox[0], t.bbox[1], page.bbox[2], t.bbox[3])
            rendered = _region_fallback(page, bbox, tol, lost, ytol, plain_only=str(broken) == "decoration")
        except Exception:
            grid = [[(c or "") for c in row] for row in t.extract()]
            rendered = table.render(table.from_grid(grid))
        prepared.append((bbox, rendered))
    boxes = [b for b, _ in prepared]

    def outside(obj):
        if "x0" not in obj or "top" not in obj:
            return True
        if id(obj) in claimed:
            return False
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return not any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in boxes)

    text_page = page.filter(outside) if boxes else page
    items = []
    for line in text_page.extract_text_lines(x_tolerance=tol, y_tolerance=ytol):
        items.append((line["top"], "text", line))
    for bbox, rendered in prepared:
        items.append((bbox[1], "table", rendered))
    for top, text in vertical:
        items.append((top, "table", text))
    items.sort(key=lambda x: x[0])

    blocks, para, last_bottom = [], [], None
    columns_at = {}  # blocks 안의 자리 -> 좌우로 나란한 묶음의 칸별 글

    def flush():
        if not para:
            return
        split = None
        if len(para) >= 2 and not rotated:
            try:
                split = _side_by_side([ln for _, ln in para], tol, ytol, text_page.chars)
            except Exception:
                split = None
        if split:
            columns_at[len(blocks)] = split
            blocks.append("")
        else:
            blocks.append("\n".join(t for t, _ in para))
        para.clear()

    for _, kind, obj in items:
        if kind == "table":
            flush()
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
            flush()
        para.append((text, obj))
        last_bottom = obj["bottom"]
    flush()
    # 잇달아 나온 나란한 묶음은 한 표의 행들로 모은다
    def width_of(block):
        rows = block.split("\n")
        if len(rows) >= 2 and rows[0].startswith("| ") and re.fullmatch(r"\|( --- \|)+", rows[1]):
            return rows[1].count("---")
        return 0

    merged, i, grown = [], 0, False  # grown: 바로 앞 블록이 나란한 묶음으로 만든(이어 붙인) 표인가
    while i < len(blocks):
        if i not in columns_at:
            block = blocks[i]
            if grown and width_of(block) and width_of(block) == width_of(merged[-1]):
                rows = block.split("\n")  # 같은 표가 선 있는 부분과 없는 부분으로 갈린 것이다
                merged[-1] += "\n" + "\n".join(rows[:1] + rows[2:])
            else:
                merged.append(block)
                grown = False
            i += 1
            continue
        rows, width = [], len(columns_at[i])
        while i in columns_at and len(columns_at[i]) == width:
            rows.append(["\n".join(_CID.sub("", c) for c in col) for col in columns_at[i]])
            i += 1
        rendered = table.render_flat(table.from_grid(rows))
        if merged and width_of(merged[-1]) == width:
            body = rendered.split("\n")
            merged[-1] += "\n" + "\n".join(body[:1] + body[2:])
        else:
            merged.append(rendered)
        grown = True
    blocks = merged
    if extra:
        blocks.append("\n".join(_CID.sub("", line) for line in extra))
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
