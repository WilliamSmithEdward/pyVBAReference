"""Build a rich description index from the MicrosoftDocs/VBA-Docs repository.

The registered Office type libraries embed almost no help text, so the
human-readable documentation (summaries, parameter descriptions, remarks,
examples, enum constant meanings, and the VBA built-in functions) is pulled
from Microsoft's published VBA reference (same content as learn.microsoft.com).

Performance design:
    * The entire reference is downloaded ONCE as a single gzipped tarball
      (~15 MB, one HTTP request) and cached on disk. Subsequent runs are fully
      offline and only re-parse the cached file.
    * Parsing streams the tarball; only the relevant ``api/*.md`` files (by
      library prefix) and the ``Language/Reference`` built-in pages are read.
      Nothing is extracted to disk.

Dependencies: Python standard library only (urllib + tarfile). ASCII output.
"""

from __future__ import annotations

import os
import re
import sys
import tarfile
import tempfile
import urllib.request

TARBALL_URL = (
    "https://codeload.github.com/MicrosoftDocs/VBA-Docs/tar.gz/refs/heads/main"
)
CACHE_PATH = os.path.join(tempfile.gettempdir(), "vba_docs_main.tar.gz")

# Matches ".../api/<ApiName>.md" inside the tarball.
_API_RE = re.compile(r"(?:^|/)api/([^/]+)\.md$")
# The VBA language reference (built-in functions / statements).
_LANG_DIR = "/language/reference/"
# Trailing disambiguator such as "(method)", "(object)", "(even)".
_TAG_RE = re.compile(r"\(([^)]*)\)\s*$")
# Inline markdown -> plain text. Link targets may contain parentheses
# (e.g. "excel.application(object).md"), so allow one nested level.
_LINK_RE = re.compile(r"\[([^\]]+)\]\((?:[^()]|\([^()]*\))*\)")
_EMPH_RE = re.compile(r"\*\*?([^*]+)\*\*?")
_WS_RE = re.compile(r"\s+")
_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
_TITLE_KIND_RE = re.compile(
    r"^(.*?)\s+(function|statement|method|property|operator|object|keyword)\b",
    re.IGNORECASE,
)

# Quality ranking deciding which file "wins" a key when several map to it.
_QUALITY = {"": 3, "method": 2, "property": 2, "object": 1}
_BUILTIN_QUALITY = {"function": 4, "method": 3, "property": 3, "statement": 2}

# Soft caps to keep output reasonable.
_REMARKS_CAP = 1200
_EXAMPLE_CAP = 1600

# Common non-ASCII punctuation -> ASCII (user preference: ASCII only).
_ASCII_MAP = {
    "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
    "\u2013": "-", "\u2014": "-", "\u2026": "...", "\u00a0": " ",
    "\u2212": "-", "\u00ab": '"', "\u00bb": '"', "\u00d7": "x",
}


# --------------------------------------------------------------------------- #
# Text helpers
# --------------------------------------------------------------------------- #

def _to_ascii(text: str) -> str:
    for bad, good in _ASCII_MAP.items():
        text = text.replace(bad, good)
    return text.encode("ascii", "ignore").decode("ascii")


def _clean(text: str) -> str:
    text = _BR_RE.sub(" ", text)
    text = _LINK_RE.sub(r"\1", text)
    text = _EMPH_RE.sub(r"\1", text)
    text = text.replace("`", "")
    text = _WS_RE.sub(" ", text).strip()
    return _to_ascii(text)


def _strip_name(cell: str) -> str:
    """Reduce a table name cell ('_Password_', '**xlCSV**') to its identifier."""
    cell = cell.strip().strip("*").strip("_").strip()
    cell = _LINK_RE.sub(r"\1", cell)
    cell = cell.strip().strip("*").strip("_").strip()
    return cell


def _strip_frontmatter(raw: str) -> tuple[dict, str]:
    """Return (frontmatter_dict, body). Only top-level scalar keys are read."""
    fm: dict[str, str] = {}
    body = raw
    if raw.startswith("---"):
        end = raw.find("\n---", 3)
        if end != -1:
            head = raw[3:end]
            nl = raw.find("\n", end + 1)
            body = raw[nl + 1:] if nl != -1 else ""
            for line in head.splitlines():
                m = re.match(r"^([A-Za-z0-9_.]+):\s*(.*)$", line)
                if m and m.group(2):
                    fm[m.group(1).lower()] = m.group(2).strip()
    return fm, body


# --------------------------------------------------------------------------- #
# Markdown table + section parsing
# --------------------------------------------------------------------------- #

def _split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    cells = re.split(r"(?<!\\)\|", line)
    return [c.replace("\\|", "|").strip() for c in cells]


def _is_sep_row(line: str) -> bool:
    s = line.strip().strip("|").strip()
    return bool(s) and set(s) <= set(" -:|")


def _tables(lines: list[str]) -> list[tuple[list[str], list[list[str]]]]:
    """Return every markdown table as (lowercase_headers, rows)."""
    out: list[tuple[list[str], list[list[str]]]] = []
    n = len(lines)
    i = 0
    while i < n - 1:
        ln = lines[i].lstrip()
        if ln.startswith("|") and _is_sep_row(lines[i + 1]):
            headers = [c.lower() for c in _split_row(lines[i])]
            rows: list[list[str]] = []
            j = i + 2
            while j < n and lines[j].lstrip().startswith("|") \
                    and not _is_sep_row(lines[j]):
                rows.append(_split_row(lines[j]))
                j += 1
            out.append((headers, rows))
            i = j
        else:
            i += 1
    return out


def _section_body(lines: list[str], name: str) -> list[str]:
    """Lines under a '## <name>...' heading up to the next '## ' heading."""
    target = name.lower()
    n = len(lines)
    i = 0
    while i < n:
        s = lines[i].strip()
        if s.startswith("## ") and s[3:].strip().lower().startswith(target):
            break
        i += 1
    if i >= n:
        return []
    i += 1
    body: list[str] = []
    while i < n and not lines[i].strip().startswith("## "):
        body.append(lines[i])
        i += 1
    return body


# --------------------------------------------------------------------------- #
# Field extractors
# --------------------------------------------------------------------------- #

def _extract_summary(lines: list[str]) -> str:
    n = len(lines)
    i = 0
    while i < n and not lines[i].lstrip().startswith("# "):
        i += 1
    i += 1
    while i < n and not lines[i].strip():
        i += 1
    para: list[str] = []
    while i < n:
        s = lines[i].strip()
        if not s:
            break
        if s.startswith(("#", "|", ">", "- ", "* ", "<!--", "[!")):
            break
        para.append(s)
        i += 1
    return _clean(" ".join(para))


def _extract_params(tables) -> dict[str, str]:
    for headers, rows in tables:
        has_desc = "description" in headers
        name_col = ("name" if "name" in headers
                    else "part" if "part" in headers else None)
        if not (has_desc and name_col) or "value" in headers:
            continue
        nc = headers.index(name_col)
        dc = headers.index("description")
        params: dict[str, str] = {}
        for r in rows:
            if len(r) <= max(nc, dc):
                continue
            name = _strip_name(r[nc])
            if name:
                params[name.lower()] = _clean(r[dc])
        if params:
            return params
    return {}


def _extract_constants(tables) -> dict[str, str]:
    for headers, rows in tables:
        if not ({"name", "value", "description"} <= set(headers)):
            continue
        nc = headers.index("name")
        dc = headers.index("description")
        consts: dict[str, str] = {}
        for r in rows:
            if len(r) <= max(nc, dc):
                continue
            name = _strip_name(r[nc])
            if name:
                consts[name.lower()] = _clean(r[dc])
        if consts:
            return consts
    return {}


def _extract_remarks(lines: list[str]) -> str:
    body = _section_body(lines, "remarks")
    if not body:
        return ""
    out: list[str] = []
    in_code = False
    for ln in body:
        s = ln.strip()
        if s.startswith("```"):
            in_code = not in_code
            continue
        if in_code or not s:
            continue
        if s.startswith(("|", "[!", "> ", "#")):
            continue
        out.append(_clean(s))
    text = " ".join(x for x in out if x)
    if len(text) > _REMARKS_CAP:
        cut = text.rfind(".", 0, _REMARKS_CAP)
        text = text[:cut + 1] if cut > 0 else text[:_REMARKS_CAP].rstrip() + "..."
    return text


def _extract_example(lines: list[str]) -> str:
    body = _section_body(lines, "example")
    if not body:
        return ""
    code: list[str] = []
    collecting = False
    for ln in body:
        if ln.strip().startswith("```"):
            if collecting:
                break
            collecting = True
            continue
        if collecting:
            code.append(ln.rstrip())
    text = _to_ascii("\n".join(code).strip())
    if len(text) > _EXAMPLE_CAP:
        text = text[:_EXAMPLE_CAP].rstrip() + "\n' ..."
    return text


# --------------------------------------------------------------------------- #
# Document model
# --------------------------------------------------------------------------- #

class DocEntry:
    __slots__ = ("summary", "params", "remarks", "example", "constants")

    def __init__(self) -> None:
        self.summary: str = ""
        self.params: dict[str, str] = {}
        self.remarks: str = ""
        self.example: str = ""
        self.constants: dict[str, str] = {}


def _parse_doc(raw: str) -> DocEntry:
    _fm, body = _strip_frontmatter(raw)
    lines = body.splitlines()
    tables = _tables(lines)
    entry = DocEntry()
    entry.summary = _extract_summary(lines)
    entry.params = _extract_params(tables)
    entry.constants = _extract_constants(tables)
    entry.remarks = _extract_remarks(lines)
    entry.example = _extract_example(lines)
    return entry


class DocIndex:
    """Case-insensitive lookup of MS Learn documentation by api name / member."""

    def __init__(self) -> None:
        self._members: dict[str, tuple[int, DocEntry]] = {}
        self._events: dict[str, tuple[int, DocEntry]] = {}
        self._builtins: dict[str, tuple[int, DocEntry]] = {}

    def __len__(self) -> int:
        return len(self._members) + len(self._events) + len(self._builtins)

    # -- population ------------------------------------------------------- #

    def _add_api(self, api_name: str, entry: DocEntry) -> None:
        if not (entry.summary or entry.params or entry.constants):
            return
        base = api_name
        tag = ""
        m = _TAG_RE.search(base)
        if m:
            tag = m.group(1).strip().lower()
            base = base[:m.start()].strip()
        key = base.lower()
        quality = _QUALITY.get(tag, 1)
        target = self._events if tag.startswith("even") else self._members
        cur = target.get(key)
        if cur is None or quality > cur[0]:
            target[key] = (quality, entry)

    def _add_builtin(self, name: str, kind: str, entry: DocEntry) -> None:
        if not entry.summary:
            return
        key = name.lower()
        quality = _BUILTIN_QUALITY.get(kind, 1)
        cur = self._builtins.get(key)
        if cur is None or quality > cur[0]:
            self._builtins[key] = (quality, entry)

    # -- lookups ---------------------------------------------------------- #

    def get_type(self, prefix: str, type_name: str) -> DocEntry | None:
        hit = self._members.get(f"{prefix}.{type_name}".lower())
        return hit[1] if hit else None

    def get_member(self, prefix: str, type_name: str, name: str
                   ) -> DocEntry | None:
        hit = self._members.get(f"{prefix}.{type_name}.{name}".lower())
        return hit[1] if hit else None

    def get_event(self, prefix: str, type_name: str, name: str
                  ) -> DocEntry | None:
        key = f"{prefix}.{type_name}.{name}".lower()
        hit = self._events.get(key) or self._members.get(key)
        return hit[1] if hit else None

    def get_builtin(self, name: str) -> DocEntry | None:
        hit = self._builtins.get(name.lower())
        return hit[1] if hit else None


# --------------------------------------------------------------------------- #
# Download + build
# --------------------------------------------------------------------------- #

def _download(url: str, dest: str) -> None:
    print("  Downloading VBA reference (MicrosoftDocs/VBA-Docs)...")
    req = urllib.request.Request(url, headers={"User-Agent": "vba-ref-scraper"})
    tmp = dest + ".part"
    with urllib.request.urlopen(req, timeout=120) as resp, open(tmp, "wb") as fh:
        total = int(resp.headers.get("Content-Length") or 0)
        got = 0
        last = -1
        while True:
            chunk = resp.read(1 << 16)
            if not chunk:
                break
            fh.write(chunk)
            got += len(chunk)
            if total:
                pct = got * 100 // total
                if pct >= last + 10:
                    last = pct - (pct % 10)
                    print(f"    {last}% ({got >> 20} MB)")
    os.replace(tmp, dest)
    print(f"  Cached to {dest} ({os.path.getsize(dest) >> 20} MB)")


def build_index(prefixes: set[str], cache_path: str = CACHE_PATH,
                force_download: bool = False) -> DocIndex:
    """Build a :class:`DocIndex` for the given lowercase library prefixes.

    A prefix of ``"vba"`` additionally pulls the VBA language built-ins
    (MsgBox, Format, Left, ...). Returns an empty index (rather than raising)
    if the reference cannot be downloaded, so callers can proceed without
    enrichment.
    """
    index = DocIndex()
    prefixes = {p.lower() for p in prefixes if p}
    if not prefixes:
        return index

    if force_download or not os.path.exists(cache_path):
        try:
            _download(TARBALL_URL, cache_path)
        except Exception as exc:  # noqa: BLE001 - network/IO best-effort
            print(f"  [warn] Could not download VBA reference: {exc}")
            print("  [warn] Continuing without MS Learn descriptions.")
            return index
    else:
        print(f"  Using cached VBA reference: {cache_path}")

    api_prefixes = tuple(f"{p}." for p in prefixes if p != "vba")
    want_builtins = "vba" in prefixes
    api_count = 0
    builtin_count = 0
    try:
        with tarfile.open(cache_path, "r:gz") as tar:
            for member in tar:
                if not member.isfile() or not member.name.endswith(".md"):
                    continue
                low = member.name.lower()
                m = _API_RE.search(member.name)
                if m and api_prefixes \
                        and m.group(1).lower().startswith(api_prefixes):
                    fh = tar.extractfile(member)
                    if fh is None:
                        continue
                    entry = _parse_doc(fh.read().decode("utf-8", "replace"))
                    index._add_api(m.group(1), entry)
                    api_count += 1
                elif want_builtins and _LANG_DIR in low:
                    fh = tar.extractfile(member)
                    if fh is None:
                        continue
                    raw = fh.read().decode("utf-8", "replace")
                    fm, _body = _strip_frontmatter(raw)
                    title = _to_ascii(fm.get("title", ""))
                    tm = _TITLE_KIND_RE.match(title)
                    if not tm:
                        continue
                    name, kind = tm.group(1).strip(), tm.group(2).lower()
                    entry = _parse_doc(raw)
                    index._add_builtin(name, kind, entry)
                    builtin_count += 1
    except (tarfile.TarError, OSError) as exc:
        print(f"  [warn] Could not read VBA reference archive: {exc}")
        return index

    detail = f"{api_count} api"
    if want_builtins:
        detail += f", {builtin_count} built-ins"
    print(f"  Indexed {detail} for {sorted(prefixes)}.")
    return index


if __name__ == "__main__":
    sel = set(sys.argv[1:]) or {"excel", "office", "vba"}
    idx = build_index(sel)
    rng = idx.get_member("excel", "Worksheet", "Protect")
    print("Worksheet.Protect:", rng.summary if rng else None)
    if rng:
        print("  params:", list(rng.params.items())[:2])
    enum = idx.get_type("excel", "XlFileFormat")
    if enum:
        print("XlFileFormat const xlCSV:", enum.constants.get("xlcsv"))
    mb = idx.get_builtin("MsgBox")
    print("MsgBox:", mb.summary if mb else None)
    if mb:
        print("  params:", list(mb.params)[:5])
