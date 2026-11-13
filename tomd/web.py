"""웹 페이지(.html, .htm) 변환."""
import re
from html.parser import HTMLParser

from . import table
from .table import Cell

_VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}
_SKIP = {"script", "style", "noscript", "template", "svg", "canvas", "iframe", "object", "head", "math"}
_BLOCK = {
    "p", "div", "section", "article", "main", "header", "footer", "nav", "aside", "form", "fieldset",
    "figure", "figcaption", "address", "details", "summary", "dl", "dt", "dd", "legend", "center", "body", "html",
}
# 새 태그가 열리면 저절로 닫히는 태그 (HTML에서는 닫는 태그를 생략할 수 있다)
_AUTO_CLOSE = {
    "p": {"p"}, "li": {"li"}, "dt": {"dt", "dd"}, "dd": {"dt", "dd"}, "tr": {"tr", "td", "th"},
    "td": {"td", "th"}, "th": {"td", "th"}, "option": {"option"}, "thead": {"tbody", "tfoot"},
    "tbody": {"thead", "tbody", "tfoot"}, "tfoot": {"thead", "tbody"},
}
_BLOCK_BREAKS_P = _BLOCK | {"ul", "ol", "table", "h1", "h2", "h3", "h4", "h5", "h6", "pre", "blockquote", "hr"}


class _Node:
    __slots__ = ("tag", "attrs", "children")

    def __init__(self, tag, attrs=None):
        self.tag = tag
        self.attrs = dict(attrs or [])
        self.children = []


class _Tree(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = _Node("#root")
        self.stack = [self.root]

    def _close_until(self, names, stop=("table", "ul", "ol", "#root")):
        for i in range(len(self.stack) - 1, 0, -1):
            tag = self.stack[i].tag
            if tag in names:
                del self.stack[i:]
                return
            if tag in stop:
                return

    def handle_starttag(self, tag, attrs):
        if tag in _AUTO_CLOSE:
            self._close_until(_AUTO_CLOSE[tag])
        if tag in _BLOCK_BREAKS_P:
            self._close_until({"p"})
        node = _Node(tag, attrs)
        self.stack[-1].children.append(node)
        if tag not in _VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(_Node(tag, attrs))

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        self.stack[-1].children.append(data)


def _hidden(node):
    if "hidden" in node.attrs or node.attrs.get("aria-hidden") == "true":
        return True
    style = (node.attrs.get("style") or "").replace(" ", "").lower()
    return "display:none" in style or "visibility:hidden" in style


class _Out:
    """블록(문단·제목·표 등) 목록과, 지금 모으고 있는 한 줄 글."""

    def __init__(self):
        self.blocks = []
        self.buf = []

    def flush(self, prefix=""):
        text = "".join(self.buf)
        self.buf.clear()
        lines = [re.sub(r"[ \t\r\f\v]+", " ", line).strip() for line in text.split("\n")]
        lines = [line for line in lines if line]
        if lines:
            self.blocks.append(prefix + "  \n".join(lines))


def _inline_text(node):
    """칸 안 글처럼 한 덩어리로 쓸 글 (블록 구분은 줄바꿈으로)."""
    out = _Out()
    _walk(node, out, 0)
    out.flush()
    return "\n".join(b.replace("  \n", "\n") for b in out.blocks)


def _table(node):
    rows = []

    def collect(n):
        for child in n.children:
            if not isinstance(child, _Node) or _hidden(child):
                continue
            if child.tag == "tr":
                row = []
                for cell in child.children:
                    if isinstance(cell, _Node) and cell.tag in ("td", "th") and not _hidden(cell):
                        def num(key):
                            try:
                                return max(1, min(1000, int(cell.attrs.get(key) or 1)))
                            except ValueError:
                                return 1
                        row.append((_inline_text(cell), num("rowspan"), num("colspan")))
                rows.append(row)
            elif child.tag in ("thead", "tbody", "tfoot"):
                collect(child)
            elif child.tag == "caption":
                continue

    collect(node)
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
    caption = next((c for c in node.children if isinstance(c, _Node) and c.tag == "caption"), None)
    rendered = table.render(cells, single="quote")
    if caption is not None:
        title = _inline_text(caption).strip()
        if title:
            rendered = f"**{title}**\n\n{rendered}" if rendered else f"**{title}**"
    return rendered


def _list(node, out, depth):
    """목록 하나를 한 블록으로 (항목 사이에 빈 줄 없이, 안쪽 목록은 들여쓰기)."""
    ordered = node.tag == "ol"
    try:
        number = int(node.attrs.get("start") or 1)
    except ValueError:
        number = 1
    lines = []
    for child in node.children:
        if not isinstance(child, _Node) or _hidden(child) or child.tag != "li":
            continue
        marker = f"{number}. " if ordered else "- "
        number += 1
        sub = _Out()
        _walk(child, sub, depth + 1)
        sub.flush()
        if not sub.blocks:
            continue
        pad = " " * len(marker)
        first = True
        for block in sub.blocks:
            for line in block.split("\n"):
                if first:
                    lines.append(marker + line)
                    first = False
                elif re.match(r"\s*(- |\d+\. )", line):
                    lines.append(pad + line)  # 안쪽 목록
                else:
                    lines.append(pad + line)
    if lines:
        out.blocks.append("\n".join(lines))


def _walk(node, out, depth):
    for child in node.children:
        if isinstance(child, str):
            # 글 속의 <b> 같은 글자가 마크다운에서 태그로 읽혀 사라지지 않게 한다
            out.buf.append(re.sub(r"<(?=[A-Za-z/!?])", "&lt;", re.sub(r"\s+", " ", child)))
            continue
        tag = child.tag
        if tag in _SKIP or _hidden(child) or tag == "title":
            continue
        if tag == "br":
            out.buf.append("\n")
        elif tag == "hr":
            out.flush()
            out.blocks.append("---")
        elif re.fullmatch(r"h[1-6]", tag):
            out.flush()
            _walk(child, out, depth)
            # 문서 제목(#)은 파일 이름이 쓰므로 한 단계 내려 쓴다
            out.flush("#" * min(6, int(tag[1]) + 1) + " ")
        elif tag in ("ul", "ol"):
            out.flush()
            _list(child, out, depth)
        elif tag == "table":
            out.flush()
            rendered = _table(child)
            if rendered:
                out.blocks.append(rendered)
        elif tag == "pre":
            out.flush()
            code = "".join(_raw_text(child)).strip("\n")
            if code.strip():
                fence = "````" if "```" in code else "```"
                out.blocks.append(f"{fence}\n{code}\n{fence}")
        elif tag == "blockquote":
            out.flush()
            sub = _Out()
            _walk(child, sub, depth)
            sub.flush()
            for block in sub.blocks:
                out.blocks.append("\n".join("> " + line for line in block.split("\n")))
        elif tag == "a":
            text = _inline_text(child).replace("\n", " ").strip()
            href = (child.attrs.get("href") or "").strip()
            if text and re.match(r"(https?|mailto):", href, re.I):
                out.buf.append(f"[{text}]({href.replace(' ', '%20')})")
            else:
                out.buf.append(text)
        elif tag in ("strong", "b", "em", "i"):
            text = _inline_text(child).replace("\n", " ").strip()
            if text:
                mark = "**" if tag in ("strong", "b") else "*"
                out.buf.append(f"{mark}{text}{mark}")
        elif tag == "code":
            text = _inline_text(child).replace("\n", " ").strip()
            if text:
                out.buf.append(f"`{text}`")
        elif tag == "img":
            alt = (child.attrs.get("alt") or "").strip()
            src = (child.attrs.get("src") or "").strip()
            if src and not src.startswith("data:"):
                out.buf.append(f"![{alt}]({src.replace(' ', '%20')})")
            elif alt:
                out.buf.append(f"[그림: {alt}]")
        elif tag == "input":
            kind = (child.attrs.get("type") or "text").lower()
            if kind in ("checkbox", "radio"):
                out.buf.append("☑ " if "checked" in child.attrs else "☐ ")
            elif kind in ("submit", "button") and child.attrs.get("value"):
                out.buf.append(f"[{child.attrs['value']}]")
        elif tag == "select":
            options = [
                _inline_text(o).strip() for o in child.children
                if isinstance(o, _Node) and o.tag == "option" and not _hidden(o)
            ]
            options = [o for o in options if o]
            if options:
                out.buf.append(" (" + " / ".join(options) + ") ")
        elif tag in ("textarea", "option"):
            continue
        elif tag in _BLOCK or tag == "li":
            out.flush()
            _walk(child, out, depth)
            out.flush()
        else:
            _walk(child, out, depth)


def _raw_text(node):
    for child in node.children:
        if isinstance(child, str):
            yield child
        elif child.tag == "br":
            yield "\n"
        elif child.tag not in _SKIP:
            yield from _raw_text(child)


def _read(path):
    data = open(path, "rb").read()
    head = data[:4096].decode("ascii", "ignore").lower()
    m = re.search(r"charset\s*=\s*[\"']?([a-z0-9_-]+)", head)
    tried = []
    if m:
        tried.append({"ks_c_5601-1987": "cp949", "euc-kr": "cp949"}.get(m.group(1), m.group(1)))
    tried += ["utf-8-sig", "cp949"]
    for enc in tried:
        try:
            return data.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="replace")


def convert_html(path):
    tree = _Tree()
    tree.feed(_read(path))
    tree.close()
    out = _Out()
    try:
        _walk(tree.root, out, 0)
        out.flush()
        text = "\n\n".join(out.blocks)
    except RecursionError:
        # 태그가 지나치게 깊게 중첩된 페이지: 구조 없이 글자만 꺼낸다
        pieces, stack = [], list(reversed(tree.root.children))
        while stack:
            item = stack.pop()
            if isinstance(item, str):
                pieces.append(item)
            elif item.tag not in _SKIP and not _hidden(item):
                if item.tag in _BLOCK_BREAKS_P or item.tag in ("br", "li", "tr", "td", "th"):
                    pieces.append("\n")
                stack.extend(reversed(item.children))
        text = "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in "".join(pieces).split("\n"))
    return re.sub(r"\n{3,}", "\n\n", text).strip()
