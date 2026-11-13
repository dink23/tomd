"""변환 작업 관리: 대기열, 단일 실행, 여러 프로세스 동시 처리, 진행률."""
import json
import datetime
import re
import os
import threading
import time
import traceback
import uuid
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from pathlib import Path

APP_NAME = "ToMD"
BIG_PDF_PAGES = 40            # 이 쪽수 이상인 PDF는 여러 프로세스로 나눠 처리한다
BIG_FILE_BYTES = 3_000_000    # 이보다 큰 다른 문서도 별도 프로세스에서
BATCH_FILES = 4               # 한 번에 이보다 많은 파일이면 여러 프로세스로 동시에
BIG_FOLDER_FILES = 300        # 폴더 안 문서가 이보다 많으면 변환 전에 물어본다
SUPPORTED = (".pdf", ".hwp", ".hwpx", ".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".docx", ".pptx", ".odt", ".html", ".htm")
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


_SKIP_DIRS = {"$recycle.bin", "system volume information", "node_modules", "__pycache__"}


def iter_supported(folder, limit=None):
    """폴더(하위 폴더 포함) 안의 변환할 수 있는 문서. 열 수 없는 폴더는 건너뛴다.

    숨김 폴더(.git 등)와 휴지통은 보지 않는다. 오피스 임시 파일(~$...)도 뺀다.
    """
    count = 0
    for root, dirs, files in os.walk(folder):  # 열 수 없는 폴더는 os.walk가 조용히 건너뛴다
        dirs[:] = sorted(d for d in dirs if not d.startswith(".") and d.lower() not in _SKIP_DIRS)
        for name in sorted(files):
            if os.path.splitext(name)[1].lower() in SUPPORTED and not name.startswith("~$"):
                count += 1
                yield Path(root) / name
                if limit is not None and count >= limit:
                    return


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
    from . import hangul, odf, office, web

    funcs = {
        ".odt": odf.convert_odt,
        ".html": web.convert_html,
        ".htm": web.convert_html,
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


def _yaml(value):
    """YAML 값 하나 (문자열은 큰따옴표로 감싸 콜론·# 같은 글자도 안전하게)."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value)
    return json.dumps(str(value), ensure_ascii=False)


def document(source, body, pages=None):
    """변환 결과 전체: 문서 정보(YAML 머리말) + 제목 + 본문.

    머리말은 옵시디언의 '속성(Properties)'으로 보이고, 다른 마크다운 도구에서는 대부분 숨겨진다.
    """
    from . import version_label

    kind = source.suffix.lower().lstrip(".")
    meta = [("title", source.stem), ("source", source.name), ("type", kind)]
    if pages:
        meta.append(("pages", pages))
    if kind == "pptx":
        meta.append(("slides", len(re.findall(r"^## 슬라이드 \d+", body, flags=re.M))))
    elif kind in ("xlsx", "xlsm", "xls"):
        sheets = len(re.findall(r"^## (?!매크로\(VBA\)$)", body, flags=re.M))
        if sheets:
            meta.append(("sheets", sheets))
        macros = len(re.findall(r"^### .+ \((?:표준 모듈|문서·클래스 모듈|모듈)\)$", body, flags=re.M))
        if macros:
            meta.append(("macros", macros))
    meta.append(("converted", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    meta.append(("converter", f"ToMD {version_label()}"))
    lines = ["---"] + [f"{key}: {_yaml(value)}" for key, value in meta] + ["tags:", "  - ToMD", "---"]
    return "\n".join(lines) + f"\n\n# {source.stem}\n\n{body}\n"


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


def _chain(exc):
    """예외와 그 원인들 (PdfminerException 안의 PDFPasswordIncorrect 같은 것까지)."""
    seen = []
    while exc is not None and exc not in seen and len(seen) < 10:
        seen.append(exc)
        inner = [a for a in getattr(exc, "args", ()) if isinstance(a, BaseException)]
        exc = exc.__cause__ or (inner[0] if inner else None) or exc.__context__
    return seen


def reason(exc):
    """실패 사유를 사용자가 알아볼 수 있는 말로 바꾼다."""
    names = {type(e).__name__ for e in _chain(exc)}
    text = " ".join(str(e) for e in _chain(exc))
    if names & {"PDFPasswordIncorrect", "PDFEncryptionError", "PDFEncryptionWarning"}:
        return "암호가 걸린 PDF입니다. 암호를 풀어 저장한 뒤 다시 변환하세요."
    if "No /Root object" in text or names & {"PDFSyntaxError", "PSEOF", "PDFNoValidXRef"}:
        return "PDF 파일이 손상되었거나 PDF 형식이 아닙니다."
    if names & {"BadZipFile", "PackageNotFoundError", "InvalidFileException"}:
        return "파일이 손상되었거나, 문서 보안(DRM)이 걸려 있거나, 확장자와 실제 형식이 다릅니다. 원래 프로그램에서 열어 다시 저장해 보세요."
    if "XLRDError" in names:
        return "엑셀(xls) 파일을 읽지 못했습니다. 손상되었거나 문서 보안(DRM)이 걸려 있을 수 있습니다. 엑셀에서 열어 xlsx로 저장한 뒤 다시 변환해 보세요."
    if "PermissionError" in names:
        return "파일을 읽거나 결과를 저장할 권한이 없습니다. 다른 프로그램에서 열려 있거나 쓰기가 막힌 폴더일 수 있습니다."
    if "MemoryError" in names:
        return "파일이 너무 커서 메모리가 부족합니다."
    if "BrokenProcessPool" in names:
        return "변환 작업이 비정상적으로 끝났습니다. 파일을 하나씩 다시 변환해 보세요."
    return str(exc) or exc.__class__.__name__


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
        self._pool = None    # 큰 PDF만 쓰는 작업 프로세스 묶음 (tomd.exe 복사본 여러 개)
        self._local = None   # 나머지는 이 프로그램 안에서 하나씩 처리한다
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
        if self._local is not None:
            self._local.shutdown(wait=False, cancel_futures=True)
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
    def _process_pool(self):
        if self._pool is None:
            import multiprocessing

            # 작업 프로세스는 늘 새로 띄운다 (Windows와 같은 방식). 스레드가 도는 중에 프로세스를
            # 복제(fork)하면 잠금이 걸린 채 복제되어 멈출 수 있다.
            self._pool = ProcessPoolExecutor(max_workers=self.workers, mp_context=multiprocessing.get_context("spawn"))
        return self._pool

    def _local_pool(self):
        if self._local is None:
            try:
                # 깊게 중첩된 문서(표 안의 표 등)를 읽을 때 스택이 모자라지 않도록 넉넉히
                threading.stack_size(64 * 1024 * 1024)
            except (ValueError, RuntimeError):
                pass
            self._local = ThreadPoolExecutor(max_workers=1, thread_name_prefix="tomd")
        return self._local

    def _fail(self, path, exc):
        with self._state_lock:
            self.failures.append((Path(path).name, reason(exc)))
        log(f"{path}: {exc!r}")

    def _expand(self, paths):
        for raw in paths:
            p = Path(raw)
            if p.is_dir():
                found = 0
                for child in iter_supported(p):
                    found += 1
                    yield child
                if not found:
                    # 아무 일도 없이 끝나면 고장으로 보이므로 이유를 알린다
                    with self._state_lock:
                        self.files_total += 1
                    self._fail(p, ValueError(
                        "이 폴더(하위 폴더 포함)에 변환할 수 있는 문서가 없습니다. 지원 형식: "
                        + ", ".join(e[1:] for e in SUPPORTED)
                    ))
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
        try:
            empty = path.stat().st_size == 0
        except OSError:
            empty = False
        if empty:
            self._fail(path, ValueError("빈 파일(0바이트)입니다."))
            return
        job = _Job(path)
        try:
            if ext == ".pdf":
                from . import pdf

                job.pages = pdf.page_count(str(path))
                if job.pages >= BIG_PDF_PAGES or self.files_total > BATCH_FILES:
                    # 큰 PDF만 여러 프로세스로 나눠 빨리 처리한다. 작은 파일까지 프로그램 복사본을
                    # 여러 개 띄우면 느려지고, 백신이 수상한 동작으로 볼 수 있다.
                    executor, workers = self._process_pool(), self.workers
                else:
                    executor, workers = self._local_pool(), 1
                step = max(4, min(25, -(-job.pages // (workers * 4))))
                for start in range(0, job.pages, step):
                    end = min(start + step, job.pages)
                    fut = executor.submit(_convert_pdf_range, str(path), start, end)
                    job.futures.append((fut, end - start))
            else:
                size = path.stat().st_size
                units = max(1, size // 200_000)
                # 큰 파일이나 여러 파일을 한꺼번에 변환할 때만 여러 프로세스로 나눈다
                heavy = size >= BIG_FILE_BYTES or self.files_total > BATCH_FILES
                executor = self._process_pool() if heavy else self._local_pool()
                job.futures.append((executor.submit(_convert_whole, str(path)), units))
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
                target = write_new(job.path, document(job.path, body, job.pages))
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
            if self._local is not None and not self._cancel.is_set():
                self._local.shutdown(wait=False)
            with self._state_lock:
                self.finished = True
