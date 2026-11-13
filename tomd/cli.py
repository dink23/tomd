"""파일을 마크다운으로 변환하는 실행 진입점.

사용법:  tomd.exe 파일또는폴더 [파일또는폴더 ...]
결과는 원본과 같은 폴더에 같은 이름의 .md 파일로 저장된다.
환경 변수 TOMD_QUIET=1 이면 창을 띄우지 않는다 (자동 빌드 검사용).
"""
import os
import sys
import time
import traceback

from . import engine, version_label
from .engine import APP_NAME, SUPPORTED


def _message_box(message, error=False):
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


def _ask(message):
    """예/아니요 질문. 창을 띄울 수 없으면 '예'로 본다."""
    if os.environ.get("TOMD_QUIET") or sys.platform != "win32":
        return True
    try:
        import ctypes

        return ctypes.windll.user32.MessageBoxW(None, message, APP_NAME, 0x4 | 0x20) == 6  # 예/아니요, 6 = 예
    except Exception:
        return True


def _confirm_folders(paths):
    """바탕화면·드라이브 전체·문서가 아주 많은 폴더는 변환 전에 확인한다 (빈 곳 오른쪽 클릭 실수 방지)."""
    home = os.path.normcase(os.path.expanduser("~"))
    for p in paths:
        if not os.path.isdir(p):
            continue
        norm = os.path.normcase(p.rstrip("\\/"))
        drive_root = os.path.splitdrive(p)[1].strip("\\/") == ""
        special = drive_root or norm == home or os.path.dirname(norm) == home and os.path.basename(norm) in (
            "desktop", "바탕 화면", "documents", "downloads", "문서", "다운로드", "onedrive"
        )
        count, started = 0, time.time()
        if not drive_root:
            for _ in engine.iter_supported(p, limit=engine.BIG_FOLDER_FILES + 1):
                count += 1
                if time.time() - started > 3:  # 세는 데 오래 걸리는 큰 폴더
                    count = engine.BIG_FOLDER_FILES + 1
                    break
        if special or count > engine.BIG_FOLDER_FILES:
            if drive_root:
                many = "드라이브 전체의"
            elif count > engine.BIG_FOLDER_FILES:
                many = f"{engine.BIG_FOLDER_FILES}개가 넘는"
            else:
                many = f"{count}개의"
            if not _ask(
                f"'{p}'\n\n이 폴더와 하위 폴더에서 변환할 문서 {many} 파일을 찾았습니다.\n"
                "모두 변환할까요? (같은 폴더에 .md 파일이 생깁니다)"
            ):
                return False
    return True


def _run_once():
    """대기열을 비울 때까지 변환하고 실패 개수를 돌려준다."""
    eng = engine.Engine()
    used_gui = False
    if not os.environ.get("TOMD_QUIET"):
        try:
            from . import gui

            gui.run(eng)
            used_gui = True
        except Exception:
            engine.log(traceback.format_exc())
    if not used_gui:
        if not eng.is_alive() and not eng.snapshot()["finished"]:
            eng.start()
        eng.join()
        snap = eng.snapshot()
        if snap["failures"]:
            lines = [f"- {n}: {m}" for n, m in snap["failures"][:8]]
            _message_box(
                f"변환하지 못한 파일이 {len(snap['failures'])}개 있습니다. (성공 {snap['ok']}개)\n\n"
                + "\n".join(lines),
                error=True,
            )
    return len(eng.snapshot()["failures"])


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        _message_box(
            "변환할 파일을 마우스 오른쪽 버튼으로 누른 뒤\n"
            "'마크다운(MD)으로 변환'을 선택하세요.\n\n"
            "지원 형식: " + ", ".join(e[1:] for e in SUPPORTED) + f"\n\nToMD {version_label()}"
        )
        return 0
    paths = [os.path.abspath(a) for a in args]
    if not _confirm_folders(paths):
        return 0
    engine.enqueue(paths)
    failed = 0
    lock = engine.Lock()
    while engine.queue_pending():
        if not lock.acquire():
            return 0  # 이미 실행 중인 창이 이어서 처리한다
        try:
            failed += _run_once()
        finally:
            lock.release()
    return 1 if failed else 0
