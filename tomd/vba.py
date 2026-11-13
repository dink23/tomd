"""엑셀 매크로(VBA) 코드 꺼내기.

매크로는 xlsm 안의 xl/vbaProject.bin(옛 xls는 파일 안의 _VBA_PROJECT_CUR)에
MS-OVBA 방식으로 압축되어 들어 있다. 코드를 실행하지 않고 글자만 꺼낸다.
"""
import io
import re
import struct
import zipfile


def _decompress(data):
    """MS-OVBA 압축 풀기 (2.4.1 CompressedContainer)."""
    if not data or data[0] != 1:
        raise ValueError("VBA 압축 형식이 아닙니다.")
    out = bytearray()
    pos = 1
    while pos + 2 <= len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        chunk_end = min(pos + size, len(data))
        pos += 2
        start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not flags & (1 << bit):
                    out.append(data[pos])
                    pos += 1
                    continue
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                done = len(out) - start
                bits = max((done - 1).bit_length(), 4)
                length_mask = 0xFFFF >> bits
                length = (token & length_mask) + 3
                offset = (token >> (16 - bits)) + 1
                src = len(out) - offset
                if src < 0:
                    raise ValueError("VBA 압축 자료가 손상되었습니다.")
                for i in range(length):
                    out.append(out[src + i])
        pos = chunk_end
    return bytes(out)


def _modules(ole, root):
    """[(모듈 이름, 종류, 코드)]  root: VBA 저장소의 경로 (예: 'VBA' 또는 '_VBA_PROJECT_CUR/VBA')."""
    d = _decompress(ole.openstream(root + "/dir").read())
    codepage = 1252
    modules, current = [], None
    pos = 0
    while pos + 6 <= len(d):
        rid, size = struct.unpack_from("<HI", d, pos)
        pos += 6
        if rid == 0x0009:  # PROJECTVERSION: 크기 칸과 달리 6바이트가 이어진다
            size = 6
        data = d[pos:pos + size]
        pos += size
        if rid == 0x0003 and len(data) >= 2:
            codepage = struct.unpack_from("<H", data)[0]
        elif rid == 0x0019:
            current = {"name": data, "uname": None, "stream": None, "offset": 0, "kind": "모듈"}
            modules.append(current)
        elif current is not None:
            if rid == 0x0047:
                current["uname"] = data
            elif rid == 0x001A:
                current["stream"] = data
            elif rid == 0x0032 and current["stream"] is not None and not isinstance(current["stream"], str):
                current["ustream"] = data
            elif rid == 0x0031 and len(data) >= 4:
                current["offset"] = struct.unpack_from("<I", data)[0]
            elif rid == 0x0022:
                current["kind"] = "문서·클래스 모듈"
            elif rid == 0x0021:
                current["kind"] = "표준 모듈"
        if rid == 0x0010:  # dir 끝
            break

    encoding = {949: "cp949", 932: "cp932", 936: "gbk", 950: "big5", 65001: "utf-8"}.get(codepage, f"cp{codepage}")

    def text(raw, unicode_raw=None):
        if unicode_raw:
            try:
                return unicode_raw.decode("utf-16-le")
            except UnicodeDecodeError:
                pass
        try:
            return raw.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            return raw.decode("cp949", errors="replace")

    out = []
    for m in modules:
        name = text(m["name"], m.get("uname"))
        stream = text(m["stream"] or m["name"], m.get("ustream"))
        path = f"{root}/{stream}"
        if not ole.exists(path):
            continue
        raw = ole.openstream(path).read()
        try:
            source = _decompress(raw[m["offset"]:])
        except Exception:
            continue
        try:
            code = source.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            code = source.decode("cp949", errors="replace")
        code = code.replace("\r\n", "\n").replace("\r", "\n")
        # 편집기 설정용 줄(Attribute VB_...)은 빼고, 실제 코드가 없는 모듈은 건너뛴다
        lines = [line for line in code.split("\n") if not line.startswith("Attribute VB_")]
        body = "\n".join(lines).strip("\n")
        meaningful = [line for line in lines if line.strip() and not line.strip().lower().startswith("option ")]
        if meaningful:
            out.append((name, m["kind"], body))
    return out


def _find_roots(ole):
    roots = []
    for entry in ole.listdir(streams=True, storages=False):
        if entry[-1].lower() == "dir" and len(entry) >= 2 and entry[-2].upper() == "VBA":
            roots.append("/".join(entry[:-1]))
    return roots


# ---------------------------------------------------------------- 민감한 정보 가리기

_MASK = "****"
_SECRET_NAME = (
    r"password|passwd|pwd|passcode|token|secret|apikey|api_key|accesskey|access_key|"
    r"privatekey|private_key|credential|비밀번호|비번|암호|패스워드"
)
_CONN = re.compile(r"(?i)\b((?:password|pwd|user\s*id|uid)\s*=\s*)(?=[^\s;\"])([^;\"\n]+)")
_REDACTIONS = [
    # 이름에 비밀번호·토큰 같은 말이 든 변수·상수·인수에 넣은 글자: pwd = "1234", Password:="1234"
    (re.compile(r'(?i)(\b(?:\w*(?:' + _SECRET_NAME + r')\w*|\w*_?pw|pass)\b\s*(?:As\s+String\s*)?:?=\s*)"([^"\n]+)"'),
     lambda m: f'{m.group(1)}"{_MASK}"'),
    # 시트·통합문서 보호 암호: ws.Unprotect "1234", .Protect "1234"
    (re.compile(r'(?i)(\.?\b(?:Unprotect|Protect)\s+)"([^"\n]+)"'), lambda m: f'{m.group(1)}"{_MASK}"'),
    # 접속 문자열 안의 계정·암호: "...;User ID=kim;Password=1234;..." (따옴표 안의 글에서만)
    (re.compile(r'"[^"\n]*"'),
     lambda m: _CONN.sub(lambda c: f"{c.group(1)}{_MASK}", m.group(0))),
    # 주소에 붙은 계정: https://kim:1234@server
    (re.compile(r"(?i)\b([a-z][a-z0-9+.-]*://)([^/\s:@\"]+):([^@\s\"/]+)@"),
     lambda m: f"{m.group(1)}{_MASK}:{_MASK}@"),
    # 개인 폴더 경로: C:\Users\홍길동\ → C:\Users\****\
    (re.compile(r"(?i)([a-z]:\\(?:users|documents and settings)\\)([^\\\"\n]+)"),
     lambda m: f"{m.group(1)}{_MASK}"),
    (re.compile(r"(/(?:Users|home)/)([^/\"\s]+)"), lambda m: f"{m.group(1)}{_MASK}"),
    # 공유 폴더 서버 이름: \\server\share → \\****\share
    (re.compile(r'(?<![\\\w])(\\\\)([^\\"\s]+)(?=\\)'), lambda m: f"{m.group(1)}{_MASK}"),
    # IP 주소
    (re.compile(r"(?<![\d.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\d.])"),
     lambda m: _MASK),
    # 이메일 주소의 아이디 부분
    (re.compile(r"\b([A-Za-z0-9._%+-]+)@([A-Za-z0-9.-]+\.[A-Za-z]{2,})\b"), lambda m: f"{_MASK}@{m.group(2)}"),
    # 주민등록번호, 휴대전화 번호
    (re.compile(r"(?<!\d)\d{6}-?[1-8]\d{6}(?!\d)"), lambda m: _MASK),
    (re.compile(r"(?<!\d)01[016789]-?\d{3,4}-?\d{4}(?!\d)"), lambda m: _MASK),
]


def redact(code):
    """비밀번호·접속 정보·개인 경로로 보이는 곳을 가린다. (가린 코드, 가린 곳 수)"""
    before = code.count(_MASK)
    for pattern, repl in _REDACTIONS:
        code = pattern.sub(repl, code)
    count = code.count(_MASK) - before
    return code, count


def extract(path):
    """파일 안의 매크로 모듈 [(이름, 종류, 코드)]. 매크로가 없으면 빈 목록."""
    import olefile

    with open(path, "rb") as f:
        head = f.read(8)
    blobs = []
    if head.startswith(b"PK"):
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if re.search(r"vbaProject\.bin$", name, re.I):
                    blobs.append(z.read(name))
    elif head.startswith(b"\xd0\xcf\x11\xe0"):
        blobs.append(path)
    modules = []
    for blob in blobs:
        source = io.BytesIO(blob) if isinstance(blob, bytes) else blob
        if not olefile.isOleFile(source):
            continue
        if isinstance(source, io.BytesIO):
            source.seek(0)
        with olefile.OleFileIO(source) as ole:
            for root in _find_roots(ole):
                modules.extend(_modules(ole, root))
    return modules


def markdown(path):
    """매크로를 마크다운 블록으로. 없거나 읽지 못하면 빈 문자열."""
    try:
        modules = extract(path)
    except Exception:
        return ""
    if not modules:
        return ""
    masked = 0
    cleaned = []
    for name, kind, code in modules:
        code, n = redact(code)
        masked += n
        cleaned.append((name, kind, code))
    note = "> 파일에 들어 있는 매크로 코드입니다. 변환 프로그램은 코드를 실행하지 않고 글자만 옮겼습니다."
    if masked:
        note += f"  \n> 비밀번호·접속 정보·개인 경로 등으로 보이는 {masked}곳은 `{_MASK}`로 가렸습니다. 원래 값은 원본 파일에서 확인하세요."
    out = ["## 매크로(VBA)", "", note]
    for name, kind, code in cleaned:
        fence = "````" if "```" in code else "```"
        out += ["", f"### {name} ({kind})", "", f"{fence}vba", code, fence]
    return "\n".join(out)
