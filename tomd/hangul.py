"""한글 문서(.hwp, .hwpx) 변환."""
import struct
import zipfile
import zlib
import re
import xml.etree.ElementTree as ET

from . import table
from .table import Cell


class ConvertError(Exception):
    pass


# ---------------------------------------------------------------- HWPX

def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _hwpx_t(node):
    """hp:t 안의 글자. 탭과 줄바꿈 요소를 반영한다."""
    text = node.text or ""
    for child in node:
        name = _local(child.tag)
        if name == "tab":
            text += "\t"
        elif name == "lineBreak":
            text += "\n"
        elif name in ("fwSpace", "nbSpace"):
            text += " "
        text += child.tail or ""
    return text


def _hwpx_paragraph(p, blocks):
    """문단 하나를 처리해 blocks에 글과 표를 차례로 넣는다."""
    text = []

    def flush():
        joined = "".join(text).strip()
        text.clear()
        if joined:
            blocks.append(joined.replace("\n", "  \n"))

    def walk(node):
        for child in node:
            name = _local(child.tag)
            if name == "t":
                text.append(_hwpx_t(child))
            elif name == "tbl":
                flush()
                rendered = _hwpx_table(child)
                if rendered:
                    blocks.append(rendered)
            elif name == "p":
                flush()
                _hwpx_paragraph(child, blocks)
            elif name in ("secPr", "linesegarray", "ctrl"):
                if name == "ctrl":
                    walk(child)
            else:
                walk(child)

    walk(p)
    flush()


def _hwpx_cell_text(tc):
    blocks = []
    for sub in tc:
        if _local(sub.tag) == "subList":
            for p in sub:
                if _local(p.tag) == "p":
                    _hwpx_paragraph(p, blocks)
    # 셀 안의 표는 글자만 남긴다
    plain = [re.sub(r"<[^>]+>", " ", b) if b.startswith("<table>") else b for b in blocks]
    return "\n".join(plain)


def _hwpx_table(tbl):
    cells = []
    for r, tr in enumerate(n for n in tbl if _local(n.tag) == "tr"):
        for c, tc in enumerate(n for n in tr if _local(n.tag) == "tc"):
            row, col, rs, cs = r, c, 1, 1
            for sub in tc:
                name = _local(sub.tag)
                if name == "cellAddr":
                    col = int(sub.get("colAddr", col))
                    row = int(sub.get("rowAddr", row))
                elif name == "cellSpan":
                    cs = int(sub.get("colSpan", 1))
                    rs = int(sub.get("rowSpan", 1))
            cells.append(Cell(row, col, _hwpx_cell_text(tc), rs, cs))
    return table.render(cells)


def convert_hwpx(path):
    blocks = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if re.fullmatch(r"Contents/section\d+\.xml", n)]
        names.sort(key=lambda n: int(re.search(r"\d+", n).group()))
        if not names:
            raise ConvertError("한글(hwpx) 문서의 본문을 찾지 못했습니다.")
        for name in names:
            root = ET.fromstring(z.read(name))
            for p in root:
                if _local(p.tag) == "p":
                    _hwpx_paragraph(p, blocks)
    return "\n\n".join(blocks)


# ---------------------------------------------------------------- HWP (5.x 바이너리)

TAG_PARA_HEADER = 66
TAG_PARA_TEXT = 67
TAG_CTRL_HEADER = 71
TAG_LIST_HEADER = 72
TAG_TABLE = 77

_CHAR_CONTROLS = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}


class _Rec:
    __slots__ = ("tag", "level", "data", "children")

    def __init__(self, tag, level, data):
        self.tag = tag
        self.level = level
        self.data = data
        self.children = []


def _records(data):
    root = _Rec(-1, -1, b"")
    stack = [root]
    pos, size_total = 0, len(data)
    while pos + 4 <= size_total:
        header = struct.unpack_from("<I", data, pos)[0]
        pos += 4
        tag = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            size = struct.unpack_from("<I", data, pos)[0]
            pos += 4
        rec = _Rec(tag, level, data[pos:pos + size])
        pos += size
        while len(stack) > 1 and stack[-1].level >= level:
            stack.pop()
        stack[-1].children.append(rec)
        stack.append(rec)
    return root


def _para_text(data):
    chars = []
    count = len(data) // 2
    i = 0
    while i < count:
        code = struct.unpack_from("<H", data, i * 2)[0]
        if code < 32:
            if code in _CHAR_CONTROLS:
                if code == 10:
                    chars.append("\n")
                elif code in (30, 31):
                    chars.append(" ")
                elif code == 24:
                    chars.append("-")
                i += 1
            else:
                if code == 9:
                    chars.append("\t")
                i += 8
            continue
        chars.append(chr(code))
        i += 1
    text = "".join(chars)
    # 짝이 맞지 않는 서러게이트 제거
    return text.encode("utf-16", "surrogatepass").decode("utf-16", "ignore").strip()


def _hwp_blocks(nodes, blocks):
    for node in nodes:
        if node.tag == TAG_PARA_HEADER:
            _hwp_paragraph(node, blocks)
        elif node.tag in (TAG_CTRL_HEADER, TAG_LIST_HEADER):
            _hwp_blocks(node.children, blocks)


def _hwp_paragraph(para, blocks):
    text = "".join(
        _para_text(c.data) for c in para.children if c.tag == TAG_PARA_TEXT
    ).strip()
    if text:
        blocks.append(text.replace("\n", "  \n"))
    for ctrl in para.children:
        if ctrl.tag != TAG_CTRL_HEADER:
            continue
        if ctrl.data[:4] == b" lbt":
            rendered = _hwp_table(ctrl)
            if rendered:
                blocks.append(rendered)
        else:
            _hwp_blocks(ctrl.children, blocks)


def _hwp_table(ctrl):
    cells, current = [], None
    order = 0
    for child in ctrl.children:
        if child.tag == TAG_LIST_HEADER:
            col = row = None
            cs = rs = 1
            if len(child.data) >= 16:
                col, row, cs, rs = struct.unpack_from("<4H", child.data, 8)
            if col is None or cs == 0 or rs == 0 or col > 512 or row > 4096:
                col, row, cs, rs = order, 0, 1, 1
            current = {"row": row, "col": col, "rs": rs, "cs": cs, "blocks": []}
            cells.append(current)
            order += 1
            _hwp_blocks(child.children, current["blocks"])
        elif child.tag == TAG_PARA_HEADER and current is not None:
            _hwp_paragraph(child, current["blocks"])
    out = []
    for c in cells:
        plain = [
            re.sub(r"<[^>]+>", " ", b) if b.startswith("<table>") else b
            for b in c["blocks"]
        ]
        out.append(Cell(c["row"], c["col"], "\n".join(plain), c["rs"], c["cs"]))
    return table.render(out)


def convert_hwp(path):
    import olefile

    if not olefile.isOleFile(path):
        with open(path, "rb") as f:
            head = f.read(4)
        if head[:2] == b"PK":
            return convert_hwpx(path)
        raise ConvertError("지원하지 않는 한글 문서 형식입니다. (한글 97 이전 형식일 수 있습니다)")
    ole = olefile.OleFileIO(path)
    try:
        if not ole.exists("FileHeader"):
            raise ConvertError("한글 문서가 아니거나 손상된 파일입니다.")
        header = ole.openstream("FileHeader").read()
        flags = struct.unpack_from("<I", header, 36)[0]
        if flags & 0x2:
            raise ConvertError("암호가 걸린 한글 문서는 변환할 수 없습니다.")
        if flags & 0x4:
            raise ConvertError("배포용 한글 문서는 변환할 수 없습니다.")
        compressed = bool(flags & 0x1)
        sections = [
            e for e in ole.listdir()
            if len(e) == 2 and e[0] == "BodyText" and e[1].startswith("Section")
        ]
        sections.sort(key=lambda e: int(e[1][7:]))
        if not sections:
            raise ConvertError("한글 문서의 본문을 찾지 못했습니다.")
        blocks = []
        for entry in sections:
            data = ole.openstream(entry).read()
            if compressed:
                data = zlib.decompress(data, -15)
            _hwp_blocks(_records(data).children, blocks)
    finally:
        ole.close()
    return "\n\n".join(blocks)
