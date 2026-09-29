"""Decide whether a security scan passed, and write its section of the report.

Every scanner in .github/workflows/security.yml reports through this script,
so there is one rule for all of them: a finding passes only if
.github/security/known-findings.toml accepts it, and an accepted entry that
matched nothing in the run fails too. Unexpected findings and stale entries
both exit 1.

    security_gate.py sarif  --tool codeql --input x.sarif --sarif-out y.sarif
    security_gate.py sarif  --tool semgrep --input x.sarif --sarif-out y.sarif
    security_gate.py yara   --input scan.ndjson --root DIR
    security_gate.py clamav --input clamscan.log --root DIR

Common options: --known FILE, --report FILE (Markdown, appended),
--meta "Label=value" (repeatable, shown in the report).

Standard library only.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import os
import sys
from dataclasses import dataclass

try:
    import tomllib
except ModuleNotFoundError:  # Python < 3.11
    tomllib = None

DEFAULT_KNOWN = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             os.pardir, "security", "known-findings.toml")

TOOL_TITLES = {"codeql": "CodeQL", "semgrep": "Semgrep", "clamav": "ClamAV",
               "yara-x": "YARA-X"}


@dataclass(frozen=True)
class Finding:
    tool: str
    rule: str
    path: str
    detail: str = ""


# --------------------------------------------------------------------------- #
# Accepted findings
# --------------------------------------------------------------------------- #

def load_known(path: str) -> list:
    if tomllib is None:
        raise SystemExit("security_gate.py needs Python 3.11+ (tomllib).")
    with open(path, "rb") as fh:
        data = tomllib.load(fh)
    entries = data.get("finding", [])
    for i, entry in enumerate(entries):
        missing = {"tool", "rule", "path", "reason"} - set(entry)
        if missing:
            raise SystemExit(f"{path}: entry {i + 1} is missing "
                             f"{sorted(missing)}")
        if entry["tool"] not in TOOL_TITLES:
            raise SystemExit(f"{path}: entry {i + 1} has unknown tool "
                             f"{entry['tool']!r}")
        if not entry["reason"].strip():
            raise SystemExit(f"{path}: entry {i + 1} has an empty reason")
    return entries


def matches(entry: dict, finding: Finding) -> bool:
    return (entry["tool"] == finding.tool and entry["rule"] == finding.rule
            and fnmatch.fnmatchcase(finding.path, entry["path"]))


def evaluate(findings: list, known: list, tool: str):
    """Split findings into (unexpected, accepted) and list stale entries."""
    entries = [e for e in known if e["tool"] == tool]
    used = set()
    unexpected, accepted = [], []
    for finding in findings:
        hit = next((i for i, e in enumerate(entries) if matches(e, finding)),
                   None)
        if hit is None:
            unexpected.append(finding)
        else:
            used.add(hit)
            accepted.append((finding, entries[hit]))
    stale = [e for i, e in enumerate(entries) if i not in used]
    return unexpected, accepted, stale


# --------------------------------------------------------------------------- #
# Scanner output parsers
# --------------------------------------------------------------------------- #

def _relative(path: str, root: str) -> str:
    path = path.replace("\\", "/")
    root = root.replace("\\", "/").rstrip("/") + "/"
    if path.startswith(root):
        path = path[len(root):]
    return path[2:] if path.startswith("./") else path


def parse_sarif(path: str, tool: str) -> list:
    """Results, plus the tool's own failures: a scanner that reports it did
    not finish, or that raised an error-level diagnostic, has not shown the
    code is clean, so those count as findings too."""
    with open(path, encoding="utf-8") as fh:
        sarif = json.load(fh)
    findings = []
    for run in sarif.get("runs", []):
        for invocation in run.get("invocations", []):
            if invocation.get("executionSuccessful") is False:
                findings.append(Finding(tool, "tool-execution-failed", "",
                                        "the scanner reported it did not "
                                        "finish"))
            for note in invocation.get("toolExecutionNotifications", []):
                if note.get("level") == "error":
                    ident = note.get("descriptor", {}).get("id", "error")
                    text = note.get("message", {}).get("text", "")
                    findings.append(Finding(tool, f"tool-error:{ident}", "",
                                            text))
        for result in run.get("results", []):
            loc = (result.get("locations") or [{}])[0].get(
                "physicalLocation", {})
            uri = loc.get("artifactLocation", {}).get("uri", "")
            line = loc.get("region", {}).get("startLine", "")
            text = result.get("message", {}).get("text", "")
            findings.append(Finding(tool, result.get("ruleId", ""),
                                    _relative(uri, ""),
                                    f"line {line}: {text}".strip()))
    return findings


def parse_yara(path: str, root: str) -> tuple:
    """Return (findings, files scanned) from `yr scan -o ndjson`."""
    findings, scanned = [], 0
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            if not line.strip():
                continue
            record = json.loads(line)
            scanned += 1
            for rule in record.get("rules", []):
                findings.append(Finding("yara-x", rule["identifier"],
                                        _relative(record["path"], root)))
    return findings, scanned


def parse_clamav(path: str, root: str) -> tuple:
    """Return (findings, files scanned) from `clamscan --infected` output."""
    findings, scanned = [], None
    with open(path, encoding="utf-8", errors="replace") as fh:
        for raw in fh:
            line = raw.rstrip("\n")
            if line.endswith(" FOUND"):
                target, signature = line[:-len(" FOUND")].rsplit(": ", 1)
                findings.append(Finding("clamav", signature,
                                        _relative(target, root)))
            elif line.startswith("Scanned files:"):
                scanned = int(line.split(":", 1)[1])
    return findings, scanned


def filter_sarif(src: str, dst: str, accepted: list) -> None:
    """Copy a SARIF file without the accepted results, for upload.

    The accepted list in the repository is the record for those; code
    scanning shows only what still needs a decision.
    """
    keys = {(f.rule, f.path) for f, _entry in accepted}
    with open(src, encoding="utf-8") as fh:
        sarif = json.load(fh)
    for run in sarif.get("runs", []):
        kept = []
        for result in run.get("results", []):
            loc = (result.get("locations") or [{}])[0].get(
                "physicalLocation", {})
            uri = _relative(loc.get("artifactLocation", {}).get("uri", ""), "")
            if (result.get("ruleId", ""), uri) not in keys:
                kept.append(result)
        run["results"] = kept
    with open(dst, "w", encoding="utf-8") as fh:
        json.dump(sarif, fh)


# --------------------------------------------------------------------------- #
# Report
# --------------------------------------------------------------------------- #

def _first_line(text: str) -> str:
    return " ".join(text.split())[:240]


def write_report(path: str, tool: str, meta: list, scanned, unexpected,
                 accepted, stale) -> None:
    ok = not unexpected and not stale
    lines = [f"## {TOOL_TITLES[tool]}", "",
             f"**Result: {'PASS' if ok else 'FAIL'}**", "",
             "| | |", "| --- | --- |"]
    for item in meta:
        label, _, value = item.partition("=")
        lines.append(f"| {label} | {value} |")
    if scanned is not None:
        lines.append(f"| Files scanned | {scanned:,} |")
    lines.append(f"| Unexpected findings | {len(unexpected)} |")
    lines.append(f"| Accepted findings | {len(accepted)} |")
    lines.append(f"| Stale accepted entries | {len(stale)} |")
    lines.append("")
    if unexpected:
        lines += ["### Unexpected findings", ""]
        for f in unexpected:
            detail = f" - {_first_line(f.detail)}" if f.detail else ""
            lines.append(f"- `{f.rule}` in `{f.path}`{detail}")
        lines.append("")
    if accepted:
        lines += ["### Accepted findings", ""]
        for f, entry in accepted:
            lines.append(f"- `{f.rule}` in `{f.path}` - "
                         f"{_first_line(entry['reason'])}")
        lines.append("")
    if stale:
        lines += ["### Stale accepted entries",
                  "",
                  "These entries matched nothing in this run. Remove them "
                  "from `.github/security/known-findings.toml`.", ""]
        for entry in stale:
            lines.append(f"- `{entry['rule']}` for `{entry['path']}`")
        lines.append("")
    with open(path, "a", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("format", choices=("sarif", "yara", "clamav"))
    parser.add_argument("--tool", choices=("codeql", "semgrep"),
                        help="which tool produced the SARIF")
    parser.add_argument("--input", required=True)
    parser.add_argument("--root", default="",
                        help="scan root, stripped from reported paths")
    parser.add_argument("--sarif-out",
                        help="write the SARIF minus accepted results here")
    parser.add_argument("--known", default=DEFAULT_KNOWN)
    parser.add_argument("--report", help="Markdown report to append to")
    parser.add_argument("--meta", action="append", default=[])
    args = parser.parse_args(argv)

    known = load_known(args.known)
    scanned = None
    if args.format == "sarif":
        if not args.tool:
            parser.error("sarif needs --tool")
        tool = args.tool
        findings = parse_sarif(args.input, tool)
    elif args.format == "yara":
        tool = "yara-x"
        findings, scanned = parse_yara(args.input, args.root)
    else:
        tool = "clamav"
        findings, scanned = parse_clamav(args.input, args.root)
        if scanned is None:
            print("::error::clamscan output has no scan summary; the scan "
                  "did not finish.")
            return 1

    unexpected, accepted, stale = evaluate(findings, known, tool)
    if args.sarif_out:
        filter_sarif(args.input, args.sarif_out, accepted)
    if args.report:
        write_report(args.report, tool, args.meta, scanned, unexpected,
                     accepted, stale)

    title = TOOL_TITLES[tool]
    for f in unexpected:
        print(f"::error file={f.path}::{title} {f.rule}: "
              f"{_first_line(f.detail) or 'unexpected finding'}")
    for entry in stale:
        print(f"::error::{title} accepted entry {entry['rule']} for "
              f"{entry['path']} matched nothing; remove it from "
              f"known-findings.toml")
    summary = (f"{title}: {len(findings)} finding(s), {len(accepted)} "
               f"accepted, {len(unexpected)} unexpected, {len(stale)} stale")
    if scanned is not None:
        summary += f", {scanned:,} files scanned"
    print(summary)
    return 1 if unexpected or stale else 0


if __name__ == "__main__":
    sys.exit(main())
