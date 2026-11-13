"""표를 마크다운(또는 병합 셀이 있으면 HTML)으로 출력한다."""
import bisect
import html


class Cell:
    __slots__ = ("row", "col", "rowspan", "colspan", "text")

    def __init__(self, row, col, text, rowspan=1, colspan=1):
        self.row = row
        self.col = col
        self.text = (text or "").strip()
        self.rowspan = max(1, rowspan)
        self.colspan = max(1, colspan)


def from_grid(grid):
    """2차원 리스트를 Cell 목록으로 바꾼다."""
    cells = []
    for r, row in enumerate(grid):
        for c, value in enumerate(row):
            cells.append(Cell(r, c, "" if value is None else str(value)))
    return cells


def _md_text(text):
    text = text.replace("\\", "\\\\").replace("|", "\\|")
    return "<br>".join(line.strip() for line in text.splitlines() if line.strip())


def _html_text(text):
    return "<br>".join(html.escape(line.strip()) for line in text.splitlines() if line.strip())


def _compact(cells):
    """글이 하나도 없는 행과 열(간격용 빈 줄·빈 칸)을 빼고, 병합 범위를 그에 맞게 줄인다."""
    rows = sorted({c.row for c in cells if c.text})
    cols = sorted({c.col for c in cells if c.text})
    out = []
    for c in cells:
        # 칸이 걸친 범위 안에 남는 행·열 (정렬된 목록에서 이분 탐색)
        r0, r1 = bisect.bisect_left(rows, c.row), bisect.bisect_left(rows, c.row + c.rowspan)
        k0, k1 = bisect.bisect_left(cols, c.col), bisect.bisect_left(cols, c.col + c.colspan)
        if r0 >= r1 or k0 >= k1:
            continue
        out.append(Cell(r0, k0, c.text, r1 - r0, k1 - k0))
    # 같은 자리로 모인 칸이 있으면 (원래 겹쳐 있던 칸) 글이 있는 쪽을 남긴다
    seen = {}
    for c in out:
        key = (c.row, c.col)
        if key not in seen or (c.text and not seen[key].text):
            seen[key] = c
    return list(seen.values())


def render(cells, single="table"):
    """Cell 목록을 표 문자열로 만든다. 내용이 없으면 빈 문자열.

    single="quote"면 칸이 하나뿐인 표(글상자)를 인용문으로 쓴다.
    """
    cells = [c for c in cells if c is not None]
    if not cells or not any(c.text for c in cells):
        return ""
    cells = _compact(cells)
    if len(cells) == 1 and single == "quote":
        lines = [line.strip() for line in cells[0].text.splitlines() if line.strip()]
        return "\n".join("> " + line + ("  " if i < len(lines) - 1 else "") for i, line in enumerate(lines))
    merged = any(c.rowspan > 1 or c.colspan > 1 for c in cells)
    if merged:
        return _render_html(cells)
    return _render_md(cells)


def _render_md(cells):
    nrows = max(c.row for c in cells) + 1
    ncols = max(c.col for c in cells) + 1
    grid = [["" for _ in range(ncols)] for _ in range(nrows)]
    for c in cells:
        grid[c.row][c.col] = _md_text(c.text)
    # 완전히 빈 행은 버린다
    grid = [row for row in grid if any(row)]
    if not grid:
        return ""
    lines = ["| " + " | ".join(grid[0]) + " |", "|" + " --- |" * ncols]
    for row in grid[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _render_html(cells):
    rows = {}
    for c in cells:
        rows.setdefault(c.row, []).append(c)
    lines = ["<table>"]
    for r in sorted(rows):
        lines.append("  <tr>")
        for c in sorted(rows[r], key=lambda x: x.col):
            attrs = ""
            if c.rowspan > 1:
                attrs += f' rowspan="{c.rowspan}"'
            if c.colspan > 1:
                attrs += f' colspan="{c.colspan}"'
            tag = "th" if r == min(rows) else "td"
            lines.append(f"    <{tag}{attrs}>{_html_text(c.text)}</{tag}>")
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines)


def render_flat(cells):
    """병합 셀을 풀어서(첫 칸에만 글을 두고) 항상 마크다운 표로 만든다.

    PDF처럼 병합 정보가 선 모양에서 추정된 것일 때 쓴다.
    """
    cells = [c for c in cells if c is not None and c.text]
    if not cells:
        return ""
    grid = {}
    for c in cells:
        key = (c.row, c.col)
        grid[key] = (grid[key] + " " + c.text) if key in grid else c.text
    rows = sorted({r for r, _ in grid})
    cols = sorted({c for _, c in grid})
    lines = []
    for i, r in enumerate(rows):
        lines.append("| " + " | ".join(_md_text(grid.get((r, c), "")) for c in cols) + " |")
        if i == 0:
            lines.append("|" + " --- |" * len(cols))
    return "\n".join(lines)
