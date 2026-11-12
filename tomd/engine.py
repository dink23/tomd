"""변환 작업 관리: 대기열, 단일 실행, 여러 프로세스 동시 처리, 진행률."""
import json
import os
import threading
import time
import traceback
import uuid
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

APP_NAME = "ToMD"
SUPPORTED = (".pdf", ".hwp", ".hwpx", ".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".docx", ".pptx")
STALE_SECONDS = 600


def base_dir():
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / APP_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


def log(message):
    try:
        with open(base_dir() / "log.txt", "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    except Exception:
        pass


# ---------------------------------------------------------------- 대기열과 잠금
# 탐색기에서 여러 파일을 선택하면 파일마다 프로그램이 따로 실행된다.
# 각 실행은 자기 파일을 대기열에 넣고, 잠금을 잡은 하나만 남아 전부 처리한다.

def _queue_dir():
    q = base_dir() / "queue"
    q.mkdir(exist_ok=True)
    return q


def enqueue(paths):
    name = f"{time.time_ns()}-{uuid.uuid4().hex}"
    tmp = _queue_dir() / (name + ".tmp")
    tmp.write_text(json.dumps([str(p) for p in paths]), encoding="utf-8")
    tmp.replace(_queue_dir() / (name + ".json"))


def take_queue():
    paths = []
    for item in sorted(_queue_dir().glob("*.json")):
        try:
            fresh = time.time() - item.stat().st_mtime < STALE_SECONDS
            data = json.loads(item.read_text(encoding="utf-8"))
            item.unlink()
            if fresh:
                paths.extend(data)
        except Exception:
            try:
                item.unlink()
            except Exception:
                pass
    return paths


def queue_pending():
    return any(_queue_dir().glob("*.json"))


class Lock:
    def __init__(self):
        self._file = None

    def acquire(self):
        f = open(base_dir() / "lock", "a+b")
        try:
            if os.name == "nt":
                import msvcrt

                f.seek(0)
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            f.close()
            return False
        self._file = f
        return True

    def release(self):
        if self._file is not None:
            try:
                self._file.close()
            finally:
                self._file = None


# ---------------------------------------------------------------- 작업 (별도 프로세스에서 실행)

def _convert_whole(path):
    from . import hangul, office

    funcs = {
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
    return funcs[Path(path).suffix.lower()](path)


def _convert_pdf_range(path, start, end):
    from . import pdf

    return pdf.convert_range(path, start, end)


def write_new(source, text):
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


class _Job:
    def __init__(self, path):
        self.path = path
        self.futures = []   # [(future, units)]
        self.pages = None   # PDF이면 전체 쪽수
        self.counted = set()


class Engine(threading.Thread):
    """대기열을 비울 때까지 변환한다. snapshot()으로 진행 상황을 읽는다."""

    def __init__(self):
        super().__init__(daemon=True)
        cpu = os.cpu_count() or 2
        self.workers = max(1, min(cpu - 1, 8))
        self._pool = None
        self._state_lock = threading.Lock()
        self._cancel = threading.Event()
        self._jobs = []
        self._seen = set()
        self.total_units = 0
        self.done_units = 0
        self.files_total = 0
        self.outputs = []
        self.failures = []   # [(파일 이름, 사유)]
        self.current = ""
        self.finished = False

    # -- 바깥에서 쓰는 것
    def snapshot(self):
        with self._state_lock:
            return {
                "total_units": self.total_units,
                "done_units": self.done_units,
                "files_total": self.files_total,
                "files_done": len(self.outputs) + len(self.failures),
                "ok": len(self.outputs),
                "failures": list(self.failures),
                "current": self.current,
                "finished": self.finished,
            }

    def cancel(self):
        self._cancel.set()
        pool = self._pool
        if pool is not None:
            procs = list((getattr(pool, "_processes", None) or {}).values())
            pool.shutdown(wait=False, cancel_futures=True)
            for p in procs:
                try:
                    p.terminate()
                except Exception:
                    pass

    # -- 내부
    def _fail(self, path, exc):
        with self._state_lock:
            self.failures.append((Path(path).name, str(exc) or exc.__class__.__name__))
        log(f"{path}: {exc}")

    def _expand(self, paths):
        for raw in paths:
            p = Path(raw)
            if p.is_dir():
                for child in sorted(p.rglob("*")):
                    if (
                        child.is_file()
                        and child.suffix.lower() in SUPPORTED
                        and not child.name.startswith("~$")
                    ):
                        yield child
            else:
                yield p

    def _add(self, path):
        key = os.path.normcase(os.path.abspath(str(path)))
        if key in self._seen:
            return
        self._seen.add(key)
        with self._state_lock:
            self.files_total += 1
            self.current = path.name
        ext = path.suffix.lower()
        if ext not in SUPPORTED:
            self._fail(path, ValueError(f"지원하지 않는 형식입니다: {ext or '(확장자 없음)'}"))
            return
        if not path.is_file():
            self._fail(path, FileNotFoundError("파일을 찾을 수 없습니다."))
            return
        if self._pool is None:
            self._pool = ProcessPoolExecutor(max_workers=self.workers)
        job = _Job(path)
        try:
            if ext == ".pdf":
                from . import pdf

                job.pages = pdf.page_count(str(path))
                step = max(4, min(25, -(-job.pages // (self.workers * 4))))
                for start in range(0, job.pages, step):
                    end = min(start + step, job.pages)
                    fut = self._pool.submit(_convert_pdf_range, str(path), start, end)
                    job.futures.append((fut, end - start))
            else:
                units = max(1, path.stat().st_size // 200_000)
                job.futures.append((self._pool.submit(_convert_whole, str(path)), units))
        except Exception as exc:
            self._fail(path, exc)
            return
        with self._state_lock:
            self.total_units += sum(u for _, u in job.futures)
        self._jobs.append(job)

    def _collect(self):
        for job in list(self._jobs):
            for fut, units in job.futures:
                if fut.done() and fut not in job.counted:
                    job.counted.add(fut)
                    with self._state_lock:
                        self.done_units += units
                        self.current = job.path.name
            if len(job.counted) < len(job.futures):
                continue
            self._jobs.remove(job)
            try:
                if job.pages is not None:
                    from . import pdf

                    parts = [item for fut, _ in job.futures for item in fut.result()]
                    body = pdf.assemble(parts, job.pages)
                else:
                    body = job.futures[0][0].result()
                body = (body or "").strip() or "> 이 파일에서 변환할 내용을 찾지 못했습니다."
                from . import version_label

                target = write_new(
                    job.path, f"# {job.path.stem}\n\n{body}\n\n<!-- ToMD {version_label()}로 변환 -->\n"
                )
                with self._state_lock:
                    self.outputs.append(target)
            except Exception as exc:
                log(traceback.format_exc())
                self._fail(job.path, exc)

    def run(self):
        idle_since = None
        try:
            while not self._cancel.is_set():
                new = take_queue()
                for path in self._expand(new):
                    if self._cancel.is_set():
                        break
                    self._add(path)
                self._collect()
                if self._jobs or new:
                    idle_since = None
                    time.sleep(0.1)
                    continue
                if idle_since is None:
                    idle_since = time.time()
                elif time.time() - idle_since > 0.8:
                    break
                time.sleep(0.1)
        except Exception:
            log(traceback.format_exc())
        finally:
            if self._pool is not None and not self._cancel.is_set():
                self._pool.shutdown(wait=False)
            with self._state_lock:
                self.finished = True
