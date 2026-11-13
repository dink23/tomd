"""원본 문서 식별과 메타데이터.

- content_hash : 원본 파일 내용의 SHA-256 (내용이 바뀌면 달라진다)
- doc_id       : 문서의 논리적 식별자. 처음 등록할 때 UUID로 만들고, 같은 문서를 수정·이동해도 유지한다.
                 해시가 같다는 이유만으로 같은 doc_id를 주지 않는다 (같은 내용의 다른 사업 문서일 수 있다).
- 문서등록 기록(registry.json): doc_id ↔ 원본 경로·해시·변환본 경로를 적어 두는 JSON 파일
- 추가 메타데이터: 폴더의 tomd-meta.json 과 명령행 --meta 로 지정 (추측해서 채우지 않는다)
"""
import datetime
import fnmatch
import hashlib
import json
import os
import re
import time
import uuid
from pathlib import Path

SCHEMA = 1
META_FILE = "tomd-meta.json"

# ToMD가 직접 계산해 쓰는 키: 바깥 메타데이터로 덮어쓸 수 없다
RESERVED = {
    "title", "source", "type", "pages", "slides", "sheets", "macros", "converted", "converter", "tags",
    "content_hash", "source_path", "source_size", "source_modified", "converter_version",
    "status", "warnings", "id_status", "id_candidates", "tomd_schema",
}
# 늘 머리말에 넣는 업무 메타데이터 (모르면 빈 값으로 둔다)
STANDARD = ("project_id", "year", "document_type")
_KEY = re.compile(r"^[A-Za-z가-힣][A-Za-z0-9_가-힣]{0,63}$")


def sha256(path, chunk=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            block = f.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def norm(path):
    return os.path.normcase(os.path.abspath(str(path)))


# ---------------------------------------------------------------- 메타데이터

def _clean_value(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list) and all(isinstance(v, (str, int, float)) for v in value):
        return [v.strip() if isinstance(v, str) else v for v in value]
    return None


def parse_cli_meta(items):
    """['project_id=S048', 'year=2027'] → (dict, 오류 목록)."""
    meta, errors = {}, []
    for item in items or []:
        if "=" not in item:
            errors.append(f"--meta 형식이 잘못되었습니다 (키=값): {item}")
            continue
        key, value = item.split("=", 1)
        key = key.strip()
        value = value.strip()
        if re.fullmatch(r"\d{4}", value) and key == "year":
            value = int(value)
        meta[key] = value
    return meta, errors


def _validate(meta, origin):
    clean, warnings = {}, []
    for key, value in (meta or {}).items():
        if not isinstance(key, str) or not _KEY.match(key):
            warnings.append(f"{origin}: 쓸 수 없는 메타데이터 이름 '{key}' (영문·한글·숫자·_ 만)")
            continue
        if key in RESERVED:
            warnings.append(f"{origin}: '{key}'는 ToMD가 직접 채우는 항목이라 지정값을 쓰지 않았습니다")
            continue
        v = _clean_value(value)
        if key == "year" and isinstance(v, str) and re.fullmatch(r"\d{4}", v):
            v = int(v)  # "2027"과 2027을 같게 (명령행 값과 맞춤)
        if v is None:
            warnings.append(f"{origin}: '{key}' 값은 글자·숫자·목록만 쓸 수 있습니다")
            continue
        clean[key] = v
    return clean, warnings


def _load_meta_file(folder):
    path = Path(folder) / META_FILE
    if not path.is_file():
        return None, []
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(data, dict):
            raise ValueError("JSON 객체가 아닙니다")
        return data, []
    except Exception as exc:
        return None, [f"{path.name}을(를) 읽지 못해 무시했습니다 ({path.parent}): {exc}"]


def collect_meta(source, cli_meta=None, levels=6):
    """한 문서의 추가 메타데이터와 경고.

    우선순위: 명령행 --meta > 문서별 지정(tomd-meta.json 의 files) > 폴더 공통값(가까운 폴더 우선) > 자동 추출(없음)
    tomd-meta.json 은 문서가 있는 폴더와 그 위로 levels 단계까지 읽는다 (먼 폴더 → 가까운 폴더 순).
    """
    source = Path(source)
    folders = [source.parent] + list(source.parent.parents)[: max(0, levels - 1)]
    folders.reverse()
    common, specific, warnings = {}, {}, []
    for folder in folders:
        data, errs = _load_meta_file(folder)
        warnings += errs
        if not data:
            continue
        files = data.get("files") or {}
        base = {k: v for k, v in data.items() if k != "files"}
        clean, w = _validate(base, f"{folder.name or folder}/{META_FILE}")
        warnings += w
        common.update(clean)  # 가까운 폴더의 값이 먼 폴더의 값을 덮는다
        if isinstance(files, dict):
            rel = source.relative_to(folder).as_posix()
            for pattern, values in files.items():
                if not isinstance(values, dict):
                    continue
                if fnmatch.fnmatchcase(rel, pattern) or fnmatch.fnmatchcase(source.name, pattern):
                    clean, w = _validate(values, f"{META_FILE} files[{pattern}]")
                    warnings += w
                    specific.update(clean)
    cli, w = _validate(cli_meta or {}, "--meta")
    warnings += w
    merged = {}
    merged.update(common)
    merged.update(specific)
    merged.update(cli)
    return merged, warnings


# ---------------------------------------------------------------- 기존 변환본 머리말 읽기

def read_front_matter(md_path, limit=80):
    """md 첫머리의 YAML 머리말에서 'key: 값' 줄만 읽는다 (목록은 무시)."""
    out = {}
    try:
        with open(md_path, encoding="utf-8") as f:
            first = f.readline()
            if first.strip() != "---":
                return out
            for _ in range(limit):
                line = f.readline()
                if not line or line.strip() == "---":
                    break
                m = re.match(r"^([A-Za-z_가-힣][\w가-힣]*):\s*(.*)$", line.rstrip("\n"))
                if not m:
                    continue
                key, raw = m.group(1), m.group(2).strip()
                if raw.startswith('"'):
                    try:
                        raw = json.loads(raw)
                    except ValueError:
                        raw = raw.strip('"')
                out[key] = raw
    except OSError:
        pass
    return out


def existing_outputs(folder, stem, names=None):
    """같은 이름으로 만든 변환본들: 이름.md, 이름 (2).md, … (names: 미리 읽어 둔 폴더 목록)"""
    folder = Path(folder)
    if names is None:
        try:
            names = os.listdir(folder)
        except OSError:
            return []
    pattern = re.compile(re.escape(stem) + r"(?: \(\d+\))?\.md$")
    return sorted((folder / n for n in names if pattern.fullmatch(n)), key=_copy_number)


def _copy_number(p):
    m = re.search(r" \((\d+)\)\.md$", p.name)
    return int(m.group(1)) if m else 1


# ---------------------------------------------------------------- 문서등록 기록

class Registry:
    """doc_id ↔ 원본·변환본 기록. 여러 실행이 함께 써도 서로 지우지 않도록, 저장할 때 다시 읽어 합친다."""

    def __init__(self, path):
        self.path = Path(path) if path else None
        self.docs = self._read()
        self._changes = {}

    def _read(self):
        if not self.path or not self.path.is_file():
            return {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            docs = data.get("documents", {})
            return docs if isinstance(docs, dict) else {}
        except Exception:
            # 깨진 기록은 지우지 않고 옆에 남긴다
            try:
                self.path.replace(self.path.with_suffix(f".broken-{int(time.time())}.json"))
            except OSError:
                pass
            return {}

    def resolve(self, source, content_hash, explicit=None, sibling_ids=()):
        """(doc_id, id_status, 후보 doc_id 목록)

        id_status:
          assigned  : 메타데이터로 지정한 doc_id
          same_path : 같은 위치의 원본으로 등록된 문서 (내용이 바뀌었어도 같은 문서)
          moved     : 같은 내용의 원본이 다른 위치에 등록되어 있었고, 그 원래 위치에는 이제 파일이 없음
          new       : 처음 보는 문서
          check     : 새 doc_id를 줬지만, 같은 문서일 수 있는 후보가 있어 확인이 필요함
        """
        key = norm(source)
        if explicit:
            return str(explicit), "assigned", []
        for doc_id, entry in self.docs.items():
            if entry.get("path_key") == key:
                pending = [d for d in entry.get("check_candidates") or [] if d in self.docs]
                if pending:  # 확인 필요로 표시한 문서는 사용자가 doc_id를 지정할 때까지 계속 표시한다
                    return doc_id, "check", pending
                return doc_id, "same_path", []
        # 기록이 없어도, 같은 자리에 있던 예전 변환본의 doc_id를 이어 받는다
        for doc_id in sibling_ids:
            if doc_id and doc_id not in self.docs:
                return doc_id, "same_path", []
        same_hash = [d for d, e in self.docs.items() if e.get("content_hash") == content_hash]
        gone = [d for d in same_hash if not os.path.exists(self.docs[d].get("path", ""))]
        if len(same_hash) == 1 and len(gone) == 1:
            return gone[0], "moved", []
        if same_hash:
            return str(uuid.uuid4()), "check", same_hash
        name = Path(source).name.lower()
        renamed = [
            d for d, e in self.docs.items()
            if Path(e.get("path", "")).name.lower() == name and not os.path.exists(e.get("path", ""))
        ]
        if renamed:
            return str(uuid.uuid4()), "check", renamed
        return str(uuid.uuid4()), "new", []

    def note(self, doc_id, source, content_hash):
        """이번 실행에서 정한 doc_id를 바로 알 수 있게 해 둔다 (같은 내용의 파일 두 개를 한꺼번에 변환할 때).
        저장은 변환에 성공해 record()를 부른 뒤에만 한다."""
        if doc_id not in self.docs:
            self.docs[doc_id] = {"path": str(Path(source).resolve()), "path_key": norm(source), "content_hash": content_hash}

    def record(self, doc_id, source, content_hash, size, output, version, meta, id_status=None, candidates=()):
        now = datetime.datetime.now().isoformat(timespec="seconds")
        entry = dict(self.docs.get(doc_id) or self._changes.get(doc_id) or {})
        history = list(entry.get("hashes") or [])
        if content_hash not in history:
            history.append(content_hash)
        outputs = list(entry.get("outputs") or [])
        if output and str(output) not in outputs:
            outputs.append(str(output))
        entry.update({
            "path": str(Path(source).resolve()),
            "path_key": norm(source),
            "content_hash": content_hash,
            "hashes": history[-50:],
            "size": size,
            "outputs": outputs[-50:],
            "converter_version": version,
            "first_seen": entry.get("first_seen") or now,
            "last_converted": now,
        })
        if id_status == "check":
            entry["check_candidates"] = list(candidates)
        else:
            entry.pop("check_candidates", None)
        for key in ("project_id", "year", "document_type"):
            if meta.get(key) not in (None, ""):
                entry[key] = meta[key]
        self._changes[doc_id] = entry
        self.docs[doc_id] = entry

    def save(self):
        if not self.path or not self._changes:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        lock = self.path.with_suffix(".lock")
        deadline = time.time() + 10
        fd = None
        while fd is None:
            try:
                fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                try:
                    if time.time() - lock.stat().st_mtime > 60:
                        lock.unlink()  # 비정상 종료로 남은 잠금
                        continue
                except OSError:
                    pass
                if time.time() > deadline:
                    raise TimeoutError("문서등록 기록이 다른 작업에 잠겨 있습니다.")
                time.sleep(0.1)
        try:
            latest = self._read()
            latest.update(self._changes)
            tmp = self.path.with_suffix(".tmp")
            tmp.write_text(
                json.dumps({"schema": SCHEMA, "documents": latest}, ensure_ascii=False, indent=1),
                encoding="utf-8",
            )
            tmp.replace(self.path)
            self.docs = latest
            self._changes = {}
        finally:
            os.close(fd)
            try:
                lock.unlink()
            except OSError:
                pass
