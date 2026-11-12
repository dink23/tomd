"""파일을 마크다운으로 변환하는 실행 진입점.

사용법:  tomd.exe 파일또는폴더 [파일또는폴더 ...]
결과는 원본과 같은 폴더에 같은 이름의 .md 파일로 저장된다.
환경 변수 TOMD_QUIET=1 이면 창을 띄우지 않는다 (자동 빌드 검사용).
"""
import os
import sys
import traceback

from . import __version__, engine
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
            "지원 형식: " + ", ".join(e[1:] for e in SUPPORTED) + f"\n\nToMD v{__version__}"
        )
        return 0
    engine.enqueue([os.path.abspath(a) for a in args])
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
