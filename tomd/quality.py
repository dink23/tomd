"""변환 품질 경고.

변환기가 남긴 안내와 원본 파일에서 바로 확인할 수 있는 사실만으로 판단한다 (추측하지 않는다).
결과는 md 머리말의 status / warnings 와 결과 JSON 에 쓰인다.
"""
import re
import zipfile


def _pages(numbers, limit=15):
    numbers = sorted(numbers)
    text = ", ".join(str(n) for n in numbers[:limit])
    return text + (f" 외 {len(numbers) - limit}쪽" if len(numbers) > limit else "")


def _pdf(parts, total):
    warnings = []
    empty = [n for n, text in parts if not (text or "").strip()]
    failed = [n for n, text in parts if "쪽을 변환하지 못했습니다" in (text or "")]
    restored = [n for n, text in parts if "글자 번호 규칙으로 복원했습니다" in (text or "")]
    lost = [n for n, text in parts if "읽지 못한 글자가" in (text or "")]
    if total and len(empty) == total:
        warnings.append("모든 쪽에서 글자를 찾지 못했습니다. 스캔한 이미지 PDF로 보입니다 (글자 인식 미지원).")
    elif empty:
        warnings.append(f"글자가 없는 쪽: {_pages(empty)} (그림·스캔 쪽일 수 있음)")
    if failed:
        warnings.append(f"변환하지 못한 쪽: {_pages(failed)}")
    if restored:
        warnings.append(f"깨진 숫자 글꼴을 복원한 쪽: {_pages(restored)} (숫자를 원본과 대조 권장)")
    if lost:
        warnings.append(f"읽지 못한 글자가 있는 쪽: {_pages(lost)}")
    return warnings


def _xlsx(path):
    """수식은 있는데 계산값이 저장되지 않은 칸(다른 프로그램이 만든 파일), 숨긴 시트."""
    warnings = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            uncached = 0
            for name in names:
                if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    data = z.read(name)
                    # <c ...><f>…</f></c> 처럼 값(<v>) 없이 끝나는 수식 칸
                    uncached += len(re.findall(rb"<c\b[^>]*>\s*<f\b[^>]*(?:/>|>[^<]*</f>)\s*</c>", data))
            if uncached:
                warnings.append(f"계산 결과가 저장되지 않은 수식 칸 {uncached}개는 빈칸으로 나옵니다 (엑셀에서 열어 저장하면 해결).")
    except (zipfile.BadZipFile, OSError, KeyError):
        pass
    return warnings


def check(path, kind, body, parts=None, total=None):
    """경고 문장 목록. 문제 없으면 빈 목록."""
    warnings = []
    if kind == "pdf" and parts is not None:
        warnings += _pdf(parts, total or len(parts))
    if kind in ("xlsx", "xlsm"):
        warnings += _xlsx(path)
    if re.search(r"^> 변경 내용 추적으로 고친 곳이 아직 수락되지 않은", body, flags=re.M):
        warnings.append("수락하지 않은 변경 내용(넣은 글·지운 글)이 있어 최종본 기준으로 옮겼습니다. 원본에서 확정 여부를 확인하세요.")
    if re.search(r"^> 이 컴퓨터에서 (워드|파워포인트)를 쓸 수 없어", body, flags=re.M):
        warnings.append("오피스 프로그램 없이 글자만 꺼냈습니다. 표·목록·서식이 원본과 다를 수 있습니다.")
    if kind in ("pptx", "ppt"):
        sections = re.split(r"^## 슬라이드 ", body, flags=re.M)[1:]
        only_pictures = 0
        for sec in sections:
            lines = [ln.strip() for ln in sec.splitlines()[1:] if ln.strip()]
            if lines and all(re.fullmatch(r"\[그림 \d+개\]", ln) for ln in lines):
                only_pictures += 1
        if only_pictures:
            warnings.append(f"그림만 있는 슬라이드 {only_pictures}장은 글로 옮길 내용이 없습니다.")
    sheets = re.search(r"^> 숨긴 시트 (\d+)개", body, flags=re.M)
    if sheets:
        warnings.append(f"숨긴 시트 {sheets.group(1)}개는 변환하지 않았습니다.")
    if re.search(r"^> 숨긴 (?:행|열) \d+개", body, flags=re.M):
        warnings.append("숨긴 행·열은 변환하지 않았습니다 (시트 아래 안내 참고).")
    masked = re.search(r"등으로 보이는 (\d+)곳은 `\*\*\*\*`로 가렸습니다", body)
    if masked:
        warnings.append(f"매크로 코드에서 비밀번호·개인 정보로 보이는 {masked.group(1)}곳을 가렸습니다.")
    # 내용이 거의 없는 경우 (머리말·안내문 제외)
    plain = re.sub(r"^>.*$|^#+ .*$|<[^>]+>|[|\-\s*`]", "", body, flags=re.M)
    if kind != "pdf" and len(plain) < 5:
        warnings.append("변환된 내용이 거의 없습니다. 원본에 글자가 있는지 확인하세요.")
    return warnings
