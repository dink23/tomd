"""예전 오피스 형식(.doc, .ppt) 변환.

1) 이 컴퓨터에 워드·파워포인트가 있으면 그 프로그램으로 새 형식(.docx/.pptx) 사본을 임시 폴더에 만들어
   기존 변환기로 바꾼다. 원본 파일은 읽기 전용으로 열고 건드리지 않는다.
2) 오피스가 없거나 실패하면 파일 안의 글자만 직접 꺼낸다 (표는 칸 구분만, 서식은 빠짐).
"""
import base64
import os
import re
import shutil
import struct
import subprocess
import tempfile
import time
from pathlib import Path

from . import table
from .hangul import ConvertError

_OFFICE_TIMEOUT = 180

# 워드: 읽기 전용으로 열어 .docx 로 저장. 이미 켜 둔 워드 창(문서)이 있으면 그 워드는 끄지 않는다.
# 암호가 걸린 문서는 엉터리 암호를 넣어 열기에 실패하게 한다 (암호 묻는 창이 떠서 멈추지 않도록).
_WORD_PS = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
$own = ($app.Documents.Count -eq 0)
try {
  $app.Visible = $false
  $app.DisplayAlerts = 0
  try { $app.AutomationSecurity = 3 } catch {}
  $doc = $app.Documents.Open($env:TOMD_IN, $false, $true, $false, 'tomd-no-password', 'tomd-no-password')
  try { $doc.SaveAs2($env:TOMD_OUT, 16) } finally { $doc.Close(0) }
} finally {
  if ($own) { $app.Quit(0) }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
"""

# 파워포인트: 창 없이 읽기 전용으로 열어 .pptx 로 저장. 파워포인트는 프로그램이 하나만 떠 있으므로
# 사용자가 열어 둔 발표 자료가 있으면 끄지 않는다. '파일::암호::' 형식으로 열어 암호 창이 뜨지 않게 한다.
_POWERPOINT_PS = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$own = ($app.Presentations.Count -eq 0)
try {
  try { $app.AutomationSecurity = 3 } catch {}
  $pres = $app.Presentations.Open($env:TOMD_IN + '::tomd-no-password::', -1, 0, 0)
  try { $pres.SaveAs($env:TOMD_OUT, 24) } finally { $pres.Close() }
} finally {
  if ($own) { $app.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
"""


class _OfficeLock:
    """여러 파일을 한꺼번에 변환할 때 오피스 프로그램을 한 번에 하나씩만 쓰게 한다 (서로 끄지 않도록)."""

    def __init__(self):
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "ToMD"
        base.mkdir(parents=True, exist_ok=True)
        self.path = base / "office.lock"
        self.fd = None

    def __enter__(self):
        deadline = time.time() + _OFFICE_TIMEOUT * 2
        while True:
            try:
                self.fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                return self
            except FileExistsError:
                try:  # 변환 중 강제 종료되어 남은 잠금은 치운다
                    if time.time() - self.path.stat().st_mtime > _OFFICE_TIMEOUT + 60:
                        self.path.unlink()
                        continue
                except OSError:
                    pass
                if time.time() > deadline:
                    raise TimeoutError
                time.sleep(0.5)

    def __exit__(self, *exc):
        try:
            os.close(self.fd)
        finally:
            try:
                self.path.unlink()
            except OSError:
                pass


def _office_resave(path, script, suffix):
    """오피스로 새 형식 사본을 만든다. 성공하면 (임시 폴더, 사본 경로), 아니면 None."""
    if os.name != "nt" or os.environ.get("TOMD_NO_OFFICE"):
        return None
    folder = tempfile.mkdtemp(prefix="tomd-")
    out = os.path.join(folder, "converted" + suffix)
    env = dict(os.environ, TOMD_IN=str(Path(path).resolve()), TOMD_OUT=out)
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    try:
        with _OfficeLock():
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-EncodedCommand", encoded],
                env=env, timeout=_OFFICE_TIMEOUT, capture_output=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
    except Exception:
        pass
    if os.path.exists(out) and os.path.getsize(out) > 0:
        return folder, out
    shutil.rmtree(folder, ignore_errors=True)
    return None


def _open_ole(path, what):
    import olefile

    if not olefile.isOleFile(path):
        with open(path, "rb") as f:
            head = f.read(4)
        if head[:2] == b"PK":
            raise ConvertError(f"확장자는 .{what}이지만 실제로는 새 형식 파일입니다. 확장자를 .{what}x 로 바꿔 변환하세요.")
        raise ConvertError(f"{'워드' if what == 'doc' else '파워포인트'} 97-2003 문서가 아니거나 손상된 파일입니다.")
    return olefile.OleFileIO(path)


# ------------------------------------------------------------------ .doc

def _doc_check(path):
    ole = _open_ole(path, "doc")
    try:
        if not ole.exists("WordDocument"):
            raise ConvertError("워드 97-2003 문서가 아니거나 손상된 파일입니다.")
        fib = ole.openstream("WordDocument").read(0x20)
        flags = struct.unpack_from("<H", fib, 0x0A)[0]
        if flags & 0x0100:
            raise ConvertError("암호가 걸린 워드 문서는 변환할 수 없습니다. 워드에서 암호를 풀고 저장한 뒤 변환하세요.")
    finally:
        ole.close()


def _sprms(grpprl):
    """문단 서식 목록에서 (sprm 번호, 값 바이트)를 차례로."""
    pos = 0
    while pos + 2 <= len(grpprl):
        sprm = struct.unpack_from("<H", grpprl, pos)[0]
        pos += 2
        spra = sprm >> 13
        if spra in (0, 1):
            size = 1
        elif spra in (2, 4, 5):
            size = 2
        elif spra == 3:
            size = 4
        elif spra == 7:
            size = 3
        else:  # 길이가 앞에 붙는 값
            if sprm == 0xD608:
                size = struct.unpack_from("<H", grpprl, pos)[0] + 1 if pos + 2 <= len(grpprl) else 0
                pos += 1
            else:
                size = grpprl[pos] if pos < len(grpprl) else 0
                pos += 1
        yield sprm, grpprl[pos:pos + size]
        pos += size


def _doc_table_runs(word, tbl, fib_lcb):
    """[(시작 위치, 끝 위치, 표 안 문단인가)] — 문단 끝 글자의 파일 위치로 찾는다."""
    fc, lcb = fib_lcb(13)  # PlcBtePapx
    plc = tbl[fc:fc + lcb]
    if lcb < 8:
        return []
    n = (len(plc) - 4) // 8
    runs = []
    for k in range(n):
        pn = struct.unpack_from("<I", plc, 4 * (n + 1) + 4 * k)[0] & 0x3FFFFF
        page = word[pn * 512:pn * 512 + 512]
        if len(page) < 512:
            continue
        crun = page[511]
        fcs = struct.unpack_from(f"<{crun + 1}I", page, 0)
        for i in range(crun):
            bx = page[4 * (crun + 1) + 13 * i]
            in_table = False
            if bx:
                at = bx * 2
                cb = page[at]
                if cb:
                    grp = page[at + 1:at + 2 * cb]
                else:
                    grp = page[at + 2:at + 2 + 2 * page[at + 1]]
                for sprm, val in _sprms(grp[2:]):
                    if sprm == 0x2416 and val and val[0]:
                        in_table = True
                    elif sprm == 0x6649 and len(val) == 4 and struct.unpack("<i", val)[0] > 0:
                        in_table = True
            runs.append((fcs[i], fcs[i + 1], in_table))
    runs.sort()
    return runs


def _doc_raw_text(path):
    """워드 97-2003 문서의 본문 글자: [(글자, 파일 위치)]와 표 안 문단 정보 (조각 표 piece table 을 따라 읽는다)."""
    ole = _open_ole(path, "doc")
    try:
        word = ole.openstream("WordDocument").read()
        flags = struct.unpack_from("<H", word, 0x0A)[0]
        table_name = "1Table" if flags & 0x0200 else "0Table"
        if not ole.exists(table_name):
            raise ConvertError("워드 문서의 본문 정보를 찾지 못했습니다.")
        tbl = ole.openstream(table_name).read()
    finally:
        ole.close()
    # FibBase(32) → csw(2) + fibRgW → cslw(2) + fibRgLw → cbRgFcLcb(2) → fibRgFcLcb (위치·길이 쌍)
    pos = 32
    csw = struct.unpack_from("<H", word, pos)[0]
    pos += 2 + csw * 2
    cslw = struct.unpack_from("<H", word, pos)[0]
    # 글자 수: 본문, 각주, 머리말·꼬리말, (예약), 메모, 미주, 글상자 — 이 순서로 이어져 있다
    ccp = struct.unpack_from("<8i", word, pos + 2 + 12)
    ccp_text = ccp[0]
    extra_ranges = []  # (이름, 시작, 끝)
    at = ccp_text
    for name, size in (("각주", ccp[1]), ("", ccp[2]), ("", ccp[3]), ("", ccp[4]), ("미주", ccp[5]), ("글상자", ccp[6])):
        if name and size > 0:
            extra_ranges.append((name, at, at + size))
        at += max(size, 0)
    limit = at
    pos += 2 + cslw * 4 + 2
    base = pos

    def fib_lcb(index):
        return struct.unpack_from("<II", word, base + index * 8)

    fc_clx, lcb_clx = fib_lcb(33)
    clx = tbl[fc_clx:fc_clx + lcb_clx]
    i = 0
    while i < len(clx) and clx[i] == 0x01:  # Prc (서식 정보) 건너뛰기
        i += 3 + struct.unpack_from("<H", clx, i + 1)[0]
    if i >= len(clx) or clx[i] != 0x02:
        raise ConvertError("워드 문서의 본문 위치 정보를 읽지 못했습니다.")
    lcb = struct.unpack_from("<I", clx, i + 1)[0]
    plc = clx[i + 5:i + 5 + lcb]
    n = (len(plc) - 4) // 12
    cps = struct.unpack_from(f"<{n + 1}i", plc, 0)
    chars = []
    for k in range(n):
        start, end = cps[k], cps[k + 1]
        if start >= limit:
            break
        end = min(end, limit)
        fc = struct.unpack_from("<I", plc, 4 * (n + 1) + 8 * k + 2)[0]
        if fc & 0x40000000:  # 1바이트 글자 (영문 코드표)
            off = (fc & 0x3FFFFFFF) // 2
            text = word[off:off + (end - start)].decode("cp1252", "replace")
            chars += [(ch, off + j, start + j) for j, ch in enumerate(text)]
        else:
            text = word[fc:fc + 2 * (end - start)].decode("utf-16-le", "replace")
            chars += [(ch, fc + 2 * j, start + j) for j, ch in enumerate(text)]
    try:
        runs = _doc_table_runs(word, tbl, fib_lcb)
    except (struct.error, IndexError):
        runs = []
    main = [(ch, fc) for ch, fc, cp in chars if cp < ccp_text]
    extras = [(name, [(ch, fc) for ch, fc, cp in chars if lo <= cp < hi]) for name, lo, hi in extra_ranges]
    return main, runs, extras


def _doc_fields(chars):
    """필드: \x13 코드 \x14 결과 \x15 → 결과만 남긴다 (결과가 없으면 지운다). 겹친 필드도 처리."""
    out, stack = [], []  # stack: 각 필드가 지금 '코드' 부분인지
    for item in chars:
        ch = item[0]
        if ch == "\x13":
            stack.append(True)
        elif ch == "\x14" and stack:
            stack[-1] = False
        elif ch == "\x15" and stack:
            stack.pop()
        elif not any(stack):
            out.append(item)
    return out


def _clean(s):
    s = s.replace("\x0b", "\n").replace("\x1e", "-").replace("\x1f", "").replace("\xa0", " ")
    s = re.sub(r"[\x00-\x08\x0c\x0e-\x1d]", "", s)
    return s


def _doc_blocks(chars, runs):
    """본문 글자를 문단·표 덩어리로 나눈다. 표 안 문단인지는 문단 서식(표 안 표시)으로 가린다.
    표의 칸은 \x07 로 끝나고, 행 끝에는 \x07 이 하나 더 붙는다."""
    import bisect

    starts = [r[0] for r in runs]

    def in_table(fc):
        i = bisect.bisect_right(starts, fc) - 1
        return i >= 0 and runs[i][0] <= fc < runs[i][1] and runs[i][2]

    chars = _doc_fields(chars)
    blocks, rows, row, cell, para = [], [], [], [], []

    def flush_table():
        nonlocal row
        if row:
            rows.append(row)
            row = []
        if rows:
            width = max(len(r) for r in rows)
            grid = [r + [""] * (width - len(r)) for r in rows]
            rendered = table.render(table.from_grid(grid), single="quote")
            if rendered:
                blocks.append(rendered)
            rows.clear()

    for ch, fc in chars:
        if ch == "\x07":
            if para or cell:
                cell.append("".join(para))
                row.append("\n".join(x for x in (_clean(c).strip() for c in cell) if x))
                cell, para = [], []
            else:  # 행 끝 표시
                if row:
                    rows.append(row)
                row = []
            continue
        if ch == "\r":
            if runs and in_table(fc):
                cell.append("".join(para))  # 칸 안의 문단 나눔
            else:
                flush_table()
                line = _clean("".join(para)).strip()
                if line:
                    blocks.append(_escape(line))
            para = []
            continue
        para.append(ch)
    flush_table()
    tail = _clean("".join(para)).strip()
    if tail:
        blocks.append(_escape(tail))
    return blocks


def _escape(line):
    line = line.replace("<", "&lt;")
    if re.match(r"(#{1,6}|>|[-*+]|\d+[.)])\s", line) or line.startswith("|"):
        line = "\\" + line  # 마크다운 문법으로 읽히지 않게
    return line


def _doc_fallback(path):
    main, runs, extras = _doc_raw_text(path)
    blocks = _doc_blocks(main, runs)
    for name, chars in extras:
        more = [b for b in _doc_blocks(chars, runs) if b.strip()]
        if more:
            blocks.append(f"### {name} 속 글" if name == "글상자" else f"### {name}")
            blocks += more
    if not blocks:
        raise ConvertError("워드 문서에서 글자를 찾지 못했습니다.")
    note = ("> 이 컴퓨터에서 워드를 쓸 수 없어 문서 안의 글자만 꺼냈습니다. 표는 칸만 나눴고 제목·목록·굵은 글씨 같은 서식은 빠졌습니다. "
            "워드가 있는 컴퓨터에서 변환하거나 워드에서 .docx 로 저장한 뒤 변환하면 더 정확합니다.")
    return "\n\n".join([note] + blocks)


def convert_doc(path):
    from . import office

    _doc_check(path)
    made = _office_resave(path, _WORD_PS, ".docx")
    if made:
        folder, copy = made
        try:
            return office.convert_docx(copy)
        finally:
            shutil.rmtree(folder, ignore_errors=True)
    return _doc_fallback(path)


# ------------------------------------------------------------------ .ppt

def _records(data, start=0, end=None):
    """파워포인트 레코드 (종류, 인스턴스, 내용 시작, 내용 끝)를 차례로 (컨테이너는 안쪽까지)."""
    end = len(data) if end is None else end
    pos = start
    while pos + 8 <= end:
        ver_inst, rtype, length = struct.unpack_from("<HHI", data, pos)
        body = pos + 8
        if body + length > end:
            length = end - body
        yield rtype, ver_inst >> 4, ver_inst & 0x0F, body, body + length
        pos = body + length


_TEXT_CHARS, _TEXT_BYTES, _CSTRING = 0x0FA0, 0x0FA8, 0x0FBA
_SLIDE, _NOTES, _MAIN_MASTER, _SLIDE_LIST, _SLIDE_PERSIST = 0x03EE, 0x03F0, 0x03F8, 0x0FF0, 0x03F3
_TEXT_HEADER = 0x0F9F


def _ppt_texts(data, start, end):
    """레코드 범위 안의 글 (글상자마다 하나씩), 문서 순서대로."""
    out = []  # (글 종류, 글) — 종류 0·6 은 제목 자리
    kind = [None]

    def walk(s, e):
        for rtype, inst, ver, b, f in _records(data, s, e):
            if ver == 0x0F:
                walk(b, f)
            elif rtype == _TEXT_HEADER and f - b >= 4:
                kind[0] = struct.unpack_from("<I", data, b)[0]
            elif rtype in (_TEXT_CHARS, _TEXT_BYTES):
                enc = "utf-16-le" if rtype == _TEXT_CHARS else "cp1252"
                out.append((kind[0], data[b:f].decode(enc, "replace")))
                kind[0] = None

    walk(start, end)
    return out


def _ppt_slides(path):
    ole = _open_ole(path, "ppt")
    try:
        if ole.exists("EncryptedSummary"):
            raise ConvertError("암호가 걸린 파워포인트 파일은 변환할 수 없습니다. 파워포인트에서 암호를 풀고 저장한 뒤 변환하세요.")
        if not ole.exists("PowerPoint Document"):
            raise ConvertError("파워포인트 97-2003 문서가 아니거나 손상된 파일입니다.")
        data = ole.openstream("PowerPoint Document").read()
        current = ole.openstream("Current User").read() if ole.exists("Current User") else b""
    finally:
        ole.close()
    # 영속 디렉터리: 객체 번호 → 스트림 안 위치 (마지막 편집부터 거슬러 올라가며 모은다)
    persist = {}
    try:
        offset = struct.unpack_from("<I", current, 16)[0]
        seen = set()
        while offset and offset not in seen and offset + 8 < len(data):
            seen.add(offset)
            # UserEditAtom: 내용 시작 +12 = offsetLastEdit, +16 = offsetPersistDirectory
            last_edit, dir_offset = struct.unpack_from("<II", data, offset + 8 + 8)
            _, _, length = struct.unpack_from("<HHI", data, dir_offset)
            pos, stop = dir_offset + 8, dir_offset + 8 + length
            while pos + 4 <= stop:
                head = struct.unpack_from("<I", data, pos)[0]
                first, count = head & 0xFFFFF, head >> 20
                pos += 4
                for k in range(count):
                    persist.setdefault(first + k, struct.unpack_from("<I", data, pos)[0])
                    pos += 4
            offset = last_edit
    except struct.error:
        persist = {}
    # 슬라이드 목록 (발표 순서)과 개요(제목·본문 자리 글)
    slides = []  # [(persist id, [개요 글])]
    notes_by_id = {}  # 노트 번호 -> (persist id, [노트 글])
    for rtype, inst, ver, b, f in _records(data):
        if rtype != 0x03E8:  # DocumentContainer
            continue
        for rt, ins, v, bb, ff in _records(data, b, f):
            if rt == _SLIDE_LIST and ins == 2:
                current = None
                for r2, i2, v2, b2, f2 in _records(data, bb, ff):
                    if r2 == _SLIDE_PERSIST and f2 - b2 >= 16:
                        pid, _, _, sid = struct.unpack_from("<IIII", data, b2)
                        current = notes_by_id.setdefault(sid, (pid, []))
                    elif r2 in (_TEXT_CHARS, _TEXT_BYTES) and current:
                        enc = "utf-16-le" if r2 == _TEXT_CHARS else "cp1252"
                        current[1].append(data[b2:f2].decode(enc, "replace"))
            if rt == _SLIDE_LIST and ins == 0:
                kind = None
                for r2, i2, v2, b2, f2 in _records(data, bb, ff):
                    if r2 == _SLIDE_PERSIST:
                        slides.append((struct.unpack_from("<I", data, b2)[0], [], []))
                    elif r2 == _TEXT_HEADER and f2 - b2 >= 4:
                        kind = struct.unpack_from("<I", data, b2)[0]
                    elif r2 in (_TEXT_CHARS, _TEXT_BYTES) and slides:
                        enc = "utf-16-le" if r2 == _TEXT_CHARS else "cp1252"
                        text = data[b2:f2].decode(enc, "replace")
                        slides[-1][1].append(text)
                        if kind in (0, 6):  # 제목·가운데 제목 자리
                            slides[-1][2].append(text)
                        kind = None
        break
    def container(pid, kind):
        at = persist.get(pid)
        if at is None or at + 8 > len(data):
            return None
        _, rtype, length = struct.unpack_from("<HHI", data, at)
        return (at + 8, at + 8 + length) if rtype == kind else None

    def slide_info(b, f):
        """(숨긴 슬라이드인가, 발표자 노트 글)"""
        hidden, notes = False, []
        for rtype, inst, ver, bb, ff in _records(data, b, f):
            if rtype == 0x03F9 and ff - bb >= 12:  # SSSlideInfoAtom
                hidden = bool(struct.unpack_from("<H", data, bb + 10)[0] & 0x0004)
            elif rtype == 0x03EF and ff - bb >= 20:  # SlideAtom: 노트 번호로 노트 슬라이드를 찾는다
                pid, listed = notes_by_id.get(struct.unpack_from("<I", data, bb + 16)[0], (None, []))
                notes = list(listed)
                where = container(pid, _NOTES) if pid is not None else None
                if where:
                    found = _ppt_texts(data, *where)
                    for t in [t for k, t in found if k == 2] or [t for k, t in found]:
                        if t not in notes:
                            notes.append(t)
        return hidden, notes

    def pick_title(found, outline_titles):
        titles = outline_titles + [t for k, t in found if k in (0, 6)]
        if titles:
            return titles[0]
        # 제목 자리가 없는 슬라이드: 맨 처음 나오는 한 줄짜리 짧은 글
        for _, t in found:
            line = _clean(t).strip()
            if line and "\r" not in t and "\n" not in line and len(line) <= 60 and re.search(r"[^\W\d_]", line):
                return t
        return None

    result = []
    for pid, outline, outline_titles in slides:
        where = container(pid, _SLIDE)
        found = _ppt_texts(data, *where) if where else []
        texts = list(outline)
        for _, t in found:
            if t not in texts:
                texts.append(t)
        hidden, notes = slide_info(*where) if where else (False, [])
        title = pick_title(found, outline_titles)
        if title in texts:
            texts.remove(title)
        result.append((title, texts, hidden, notes))
    if not result:  # 목록을 못 읽은 파일: 슬라이드 레코드를 나온 차례대로
        for rtype, inst, ver, b, f in _records(data):
            if rtype == _SLIDE:
                found = _ppt_texts(data, b, f)
                title = pick_title(found, [])
                hidden, notes = slide_info(b, f)
                result.append((title, [t for _, t in found if t != title], hidden, notes))
    return result


def _ppt_fallback(path):
    slides = _ppt_slides(path)
    if not any(title or texts for title, texts, _, _ in slides):
        raise ConvertError("파워포인트 파일에서 글자를 찾지 못했습니다.")
    parts = ["> 이 컴퓨터에서 파워포인트를 쓸 수 없어 슬라이드 안의 글자만 꺼냈습니다. 표·목록·서식은 빠졌습니다. "
             "파워포인트가 있는 컴퓨터에서 변환하거나 파워포인트에서 .pptx 로 저장한 뒤 변환하면 더 정확합니다."]
    def lines_of(t):
        lines = [_escape(_clean(x).strip()) for x in t.replace("\r", "\n").split("\n")]
        return [x for x in lines if x and not re.fullmatch(r"\*", x)]

    for n, (title, texts, hidden, notes) in enumerate(slides, 1):
        title = re.sub(r"\s+", " ", _clean(title or "")).strip()
        parts.append(f"## 슬라이드 {n}" + (" (숨김)" if hidden else "") + (f": {title}" if title else ""))
        body = []
        for t in texts:
            lines = lines_of(t)
            if lines:
                body.append("\n".join(f"{x}  " if k < len(lines) - 1 else x for k, x in enumerate(lines)))
        note_lines = [x for t in notes for x in lines_of(t)]
        if note_lines:
            body.append("> 발표자 노트: " + "\n> ".join(note_lines))
        parts.append("\n\n".join(body) if body else "(글 없음)")
    return "\n\n".join(parts)


def convert_ppt(path):
    from . import office

    _ppt_slides(path)  # 암호·손상 여부를 먼저 확인한다
    made = _office_resave(path, _POWERPOINT_PS, ".pptx")
    if made:
        folder, copy = made
        try:
            return office.convert_pptx(copy)
        finally:
            shutil.rmtree(folder, ignore_errors=True)
    return _ppt_fallback(path)
