"""파일을 마크다운으로 변환하는 실행 진입점.

사용법:  tomd.exe 파일또는폴더 [파일또는폴더 ...]
결과는 원본과 같은 폴더에 같은 이름의 .md 파일로 저장된다.
"""
import datetime
import os
import sys
import traceback
from pathlib import Path

APP_NAME = "ToMD"


def _converters():
    from . import hangul, office, pdf

    return {
        ".pdf": pdf.convert_pdf,
        ".hwp": hangul.convert_hwp,
        ".hwpx": hangul.convert_hwpx,
        ".xlsx": office.convert_xlsx,
        ".xlsm": office.convert_xlsx,
        ".xls": office.convert_xls,
        ".csv": office.convert_csv,
        ".tsv": office.convert_csv,
        ".docx": office.convert_docx,
        ".pptx": office.convert_pptx,
    }


SUPPORTED = (".pdf", ".hwp", ".hwpx", ".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".docx", ".pptx")


def _write_new(source, text):
    """기존 파일을 덮어쓰지 않고 새 .md 파일을 만든다."""
    n = 1
    while True:
        suffix = "" if n == 1 else f" ({n})"
        target = source.with_name(f"{source.stem}{suffix}.md")
        try:
            with open(target, "x", encoding="utf-8", newline="\n") as f:
                f.write(text)
            return target
        except FileExistsError:
            n += 1


def convert_file(source):
    source = Path(source)
    func = _converters().get(source.suffix.lower())
    if func is None:
        raise ValueError(f"지원하지 않는 형식입니다: {source.suffix or '(확장자 없음)'}")
    body = (func(str(source)) or "").strip()
    if not body:
        body = "> 이 파일에서 변환할 내용을 찾지 못했습니다."
    text = f"# {source.stem}\n\n{body}\n"
    return _write_new(source, text)


def _collect(args):
    files = []
    for arg in args:
        p = Path(arg)
        if p.is_dir():
            for child in sorted(p.rglob("*")):
                if (
                    child.is_file()
                    and child.suffix.lower() in SUPPORTED
                    and not child.name.startswith("~$")
                ):
                    files.append(child)
        else:
            files.append(p)
    return files


def _log(message):
    try:
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / APP_NAME
        base.mkdir(parents=True, exist_ok=True)
        with open(base / "log.txt", "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {message}\n")
    except Exception:
        pass


def _notify(message, error=False):
    if os.environ.get("TOMD_QUIET"):
        return
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(None, message, APP_NAME, 0x10 if error else 0x40)
            return
        except Exception:
            pass
    stream = sys.stderr if error else sys.stdout
    if stream is not None:
        print(message, file=stream)


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        _notify(
            "변환할 파일을 마우스 오른쪽 버튼으로 누른 뒤\n"
            "'마크다운(MD)으로 변환'을 선택하세요.\n\n"
            "지원 형식: " + ", ".join(e[1:] for e in SUPPORTED)
        )
        return 0
    from_folder = any(Path(a).is_dir() for a in args)
    files = _collect(args)
    done, failed = [], []
    for f in files:
        try:
            done.append(convert_file(f))
        except Exception as exc:
            failed.append((f, exc))
            _log(f"{f}: {exc}\n{traceback.format_exc()}")
    if failed:
        lines = [f"- {f.name}: {exc}" for f, exc in failed[:8]]
        if len(failed) > 8:
            lines.append(f"- 그 밖에 {len(failed) - 8}개")
        _notify(
            f"변환하지 못한 파일이 {len(failed)}개 있습니다. (성공 {len(done)}개)\n\n" + "\n".join(lines),
            error=True,
        )
        return 1
    if from_folder:
        _notify(f"{len(done)}개 파일을 변환했습니다." if done else "폴더에 변환할 수 있는 파일이 없습니다.")
    elif sys.stdout is not None:
        for d in done:
            print(d)
    return 0
