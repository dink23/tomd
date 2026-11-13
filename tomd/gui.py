"""변환 진행 창."""
import time

SHOW_AFTER = 0.6    # 이 시간 안에 끝나는 작업은 창을 띄우지 않는다
CLOSE_AFTER = 1500  # 성공 후 창을 닫기까지 (밀리초)


def run(engine):
    """진행 창을 띄우고 engine이 끝날 때까지 기다린다."""
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.withdraw()
    from . import version_label

    root.title(f"ToMD {version_label()} - 마크다운 변환")
    root.resizable(False, False)
    font = ("맑은 고딕", 10)

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)
    title = ttk.Label(frame, text="변환을 준비하고 있습니다", font=("맑은 고딕", 11, "bold"))
    title.pack(anchor="w")
    detail = ttk.Label(frame, text="", font=font, width=56)
    detail.pack(anchor="w", pady=(6, 8))
    bar = ttk.Progressbar(frame, length=420, maximum=1000, mode="determinate")
    bar.pack(fill="x")
    percent = ttk.Label(frame, text="", font=font)
    percent.pack(anchor="e", pady=(4, 0))
    errors = tk.Text(frame, height=8, width=58, font=font, wrap="word", relief="solid", borderwidth=1)
    button = ttk.Button(frame, text="취소")
    button.pack(anchor="e", pady=(10, 0))

    state = {"shown": False, "started": time.time(), "done": False}

    def show():
        if state["shown"]:
            return
        state["shown"] = True
        root.update_idletasks()
        w, h = root.winfo_reqwidth(), root.winfo_reqheight()
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 3
        root.geometry(f"+{x}+{y}")
        root.deiconify()
        root.attributes("-topmost", True)
        root.after(800, lambda: root.attributes("-topmost", False))

    def close():
        if not state["done"]:
            engine.cancel()
        root.destroy()

    button.configure(command=close)
    root.protocol("WM_DELETE_WINDOW", close)

    def shorten(name, limit=44):
        return name if len(name) <= limit else name[: limit - 1] + "…"

    def tick():
        try:
            _tick()
        except Exception:
            import traceback

            from .engine import log

            log(traceback.format_exc())
            close()

    def _tick():
        snap = engine.snapshot()
        total = max(snap["total_units"], 1)
        ratio = min(snap["done_units"] / total, 1.0)
        if snap["finished"]:
            state["done"] = True
            bar["value"] = 1000
            percent.configure(text="100%")
            button.configure(text="닫기")
            if snap["failures"]:
                title.configure(text=f"변환하지 못한 파일이 {len(snap['failures'])}개 있습니다")
                skipped = f", 이미 변환되어 건너뜀 {len(snap['skipped'])}개" if snap["skipped"] else ""
                detail.configure(text=f"성공 {snap['ok']}개, 실패 {len(snap['failures'])}개{skipped}")
                errors.insert("1.0", "\n".join(f"· {n}: {m}" for n, m in snap["failures"]))
                errors.configure(state="disabled")
                errors.pack(fill="x", pady=(8, 0), before=button)
                show()
            elif snap["skipped"] and not snap["ok"]:
                # 아무것도 만들지 않았으면 이유를 보여 준다 (조용히 끝나면 고장으로 보인다)
                title.configure(text="이미 변환한 파일입니다")
                detail.configure(text=f"같은 내용을 같은 버전으로 변환한 .md가 있어 {len(snap['skipped'])}개를 건너뛰었습니다.")
                show()
            elif state["shown"]:
                title.configure(text="변환을 마쳤습니다")
                skipped = f" (이미 변환된 {len(snap['skipped'])}개는 건너뜀)" if snap["skipped"] else ""
                detail.configure(text=f"{snap['ok']}개 파일을 변환했습니다.{skipped}")
                root.after(CLOSE_AFTER, root.destroy)
            else:
                root.destroy()
            return
        if snap["files_total"]:
            title.configure(text="마크다운으로 변환하고 있습니다")
            detail.configure(
                text=f"{snap['files_done']}/{snap['files_total']}개 완료 · {shorten(snap['current'])}"
            )
            bar["value"] = int(ratio * 1000)
            percent.configure(text=f"{int(ratio * 100)}%")
        if time.time() - state["started"] > SHOW_AFTER:
            show()
        root.after(150, tick)

    engine.start()
    root.after(100, tick)
    root.mainloop()
