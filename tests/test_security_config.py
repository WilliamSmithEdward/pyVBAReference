"""Tests for the security scanning configuration and the gate that enforces it.

They guard the properties the Security workflow relies on: every pin stays a
pin, the accepted-findings list stays valid and documented, and the gate
decides pass and fail the way SECURITY.md says it does.
"""

from __future__ import annotations

import importlib.util
import json
import re
import sys
from pathlib import Path

import pytest

tomllib = pytest.importorskip("tomllib")

ROOT = Path(__file__).resolve().parent.parent
GITHUB = ROOT / ".github"
if not GITHUB.is_dir():
    pytest.skip("the .github folder is not part of this checkout",
                allow_module_level=True)

HEX64 = re.compile(r"^[0-9a-f]{64}$")


def _gate():
    spec = importlib.util.spec_from_file_location(
        "security_gate", GITHUB / "scripts" / "security_gate.py")
    module = importlib.util.module_from_spec(spec)
    # Dataclasses look their module up in sys.modules while being built.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


gate = _gate()


# --------------------------------------------------------------------------- #
# Pins
# --------------------------------------------------------------------------- #

def _workflows():
    return sorted((GITHUB / "workflows").glob("*.yml"))


def test_every_action_is_pinned_to_a_commit_sha():
    uses = re.compile(r"^\s*-?\s*uses:\s*(\S+)(.*)$")
    checked = 0
    for workflow in _workflows():
        for number, line in enumerate(workflow.read_text("utf-8").splitlines(), 1):
            m = uses.match(line)
            if not m or m.group(1).startswith("./"):
                continue
            ref = m.group(1).rsplit("@", 1)
            where = f"{workflow.name}:{number}"
            assert len(ref) == 2 and re.fullmatch(r"[0-9a-f]{40}", ref[1]), \
                f"{where} is not pinned to a commit SHA: {m.group(1)}"
            assert re.search(r"#\s*v\d", m.group(2)), \
                f"{where} has no version comment after its SHA"
            checked += 1
    assert checked > 0


def test_runners_are_named_releases():
    for workflow in _workflows():
        for line in workflow.read_text("utf-8").splitlines():
            if "runs-on:" in line:
                assert "latest" not in line, f"{workflow.name}: {line.strip()}"


def test_container_images_are_pinned_by_digest():
    dockerfiles = sorted((GITHUB / "security").glob("*/Dockerfile"))
    assert dockerfiles
    for dockerfile in dockerfiles:
        froms = [line for line in dockerfile.read_text("utf-8").splitlines()
                 if line.startswith("FROM ")]
        assert len(froms) == 1, dockerfile
        assert re.search(r":[\w.-]+@sha256:[0-9a-f]{64}$", froms[0]), \
            f"{dockerfile} is not pinned by tag and digest: {froms[0]}"


def test_ci_requirements_are_exact_and_hashed():
    locks = sorted((GITHUB / "requirements").glob("*.txt"))
    assert locks
    for lock in locks:
        text = lock.read_text("utf-8")
        requirements = re.findall(r"^([A-Za-z0-9]\S*)", text, re.MULTILINE)
        hashed = re.findall(
            r"^([A-Za-z0-9]\S*)[^\n]* \\\n\s+--hash=sha256:[0-9a-f]{64}",
            text, re.MULTILINE)
        assert requirements, lock.name
        for requirement in requirements:
            assert "==" in requirement, f"{lock.name}: not an exact pin: {requirement}"
        assert hashed == requirements, f"{lock.name}: every requirement needs a hash"


def test_yara_pins_are_ones_the_standard_updater_accepts():
    """.github/security/yara.json has the shape the standard updater
    (.github/security/yara_update.py, tested in repo-standards) reads and writes."""
    import importlib.util
    import json

    spec = importlib.util.spec_from_file_location("yara_update", GITHUB / "security" / "yara_update.py")
    updater = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(updater)
    pins = json.loads((GITHUB / "security" / "yara.json").read_text("utf-8"))
    assert set(pins) == {"yara_forge", "yara_x"}
    updater.check_move(pins, pins)


def _jobs(workflow: Path) -> dict:
    """Top-level job ids of a workflow mapped to their source blocks."""
    text = workflow.read_text("utf-8")
    body = text.split("\njobs:\n", 1)[1]
    parts = re.split(r"^  ([A-Za-z][\w-]*):\s*$", body, flags=re.MULTILINE)
    return dict(zip(parts[1::2], parts[2::2]))


@pytest.mark.parametrize("workflow, gate_id, gate_name", [
    ("security.yml", "security-passed", "Security passed"),
    ("malware-scan.yml", "malware-scan-passed", "Malware scan passed"),
])
def test_gate_job_needs_every_job_and_cannot_be_skipped(workflow, gate_id, gate_name):
    # The ruleset requires only the gate jobs. Each must depend on every
    # other job in its workflow, or a failing job could merge, and it must
    # run on always(), because a skipped required check counts as passing.
    jobs = _jobs(GITHUB / "workflows" / workflow)
    gate_job = jobs.pop(gate_id)
    assert f"name: {gate_name}\n" in gate_job
    assert "if: always()" in gate_job
    needs = re.search(r"needs: \[([^\]]*)\]", gate_job).group(1)
    assert sorted(n.strip() for n in needs.split(",")) == sorted(jobs)


# --------------------------------------------------------------------------- #
# Accepted findings
# --------------------------------------------------------------------------- #

def test_known_findings_are_valid_and_documented():
    entries = gate.load_known(str(GITHUB / "security" / "known-findings.toml"))
    security_md = (ROOT / "SECURITY.md").read_text("utf-8")
    for entry in entries:
        short = entry["rule"].rsplit(".", 1)[-1]
        assert f"`{short}`" in security_md, \
            f"SECURITY.md does not list the accepted {entry['rule']}"
        assert len(entry["reason"].split()) >= 10, \
            f"{entry['rule']}: the reason is too thin to re-check later"


# --------------------------------------------------------------------------- #
# The gate's decisions
# --------------------------------------------------------------------------- #

KNOWN = [
    {"tool": "yara-x", "rule": "R1", "path": "*docs/a.md", "reason": "ok"},
    {"tool": "yara-x", "rule": "R2", "path": "source/gone.txt",
     "reason": "ok"},
    {"tool": "clamav", "rule": "R1", "path": "*docs/a.md", "reason": "ok"},
]


def test_accepted_finding_passes_and_glob_crosses_directories():
    findings = [gate.Finding("yara-x", "R1", "source/docs/a.md"),
                gate.Finding("yara-x", "R1",
                             "unpacked/sdist/pkg-1.0/docs/a.md")]
    unexpected, accepted, stale = gate.evaluate(findings, KNOWN[:1], "yara-x")
    assert unexpected == [] and len(accepted) == 2 and stale == []


def test_unlisted_finding_fails():
    findings = [gate.Finding("yara-x", "R9", "source/docs/a.md")]
    unexpected, _accepted, _stale = gate.evaluate(findings, KNOWN, "yara-x")
    assert unexpected == findings


def test_entries_only_accept_their_own_tool():
    # KNOWN accepts R1 for clamav too, but a yara-x R1 elsewhere is new.
    findings = [gate.Finding("yara-x", "R1", "source/other.md")]
    unexpected, _accepted, _stale = gate.evaluate(findings, KNOWN, "yara-x")
    assert unexpected == findings


def test_entry_matching_nothing_is_stale():
    findings = [gate.Finding("yara-x", "R1", "source/docs/a.md")]
    _unexpected, _accepted, stale = gate.evaluate(findings, KNOWN, "yara-x")
    assert [e["rule"] for e in stale] == ["R2"]


def test_main_fails_on_stale_entries(tmp_path):
    known = tmp_path / "known.toml"
    known.write_text('[[finding]]\ntool = "yara-x"\nrule = "R1"\n'
                     'path = "x"\nreason = "accepted for the test"\n')
    scan = tmp_path / "scan.ndjson"
    scan.write_text(json.dumps({"path": "/s/x", "rules": []}) + "\n")
    rc = gate.main(["yara", "--input", str(scan), "--root", "/s",
                    "--known", str(known)])
    assert rc == 1


def test_parse_clamav(tmp_path):
    log = tmp_path / "clamscan.log"
    log.write_text("/scan/source/a b/x.com: Eicar-Test-Signature FOUND\n"
                   "\n----------- SCAN SUMMARY -----------\n"
                   "Scanned files: 12\nInfected files: 1\n")
    findings, scanned = gate.parse_clamav(str(log), "/scan")
    assert scanned == 12
    assert findings == [gate.Finding("clamav", "Eicar-Test-Signature",
                                     "source/a b/x.com")]


def test_clamav_without_summary_fails(tmp_path):
    # A scan that died before its summary found nothing because it did not
    # finish, which must not read as clean.
    log = tmp_path / "clamscan.log"
    log.write_text("LibClamAV Error: something broke\n")
    known = tmp_path / "known.toml"
    known.write_text("")
    rc = gate.main(["clamav", "--input", str(log), "--root", "/scan",
                    "--known", str(known)])
    assert rc == 1


def test_parse_yara_counts_files_and_relativizes_paths(tmp_path):
    ndjson = tmp_path / "scan.ndjson"
    ndjson.write_text("\n".join([
        json.dumps({"path": "/tmp/scan/source/a.md",
                    "rules": [{"identifier": "R1"}]}),
        json.dumps({"path": "/tmp/scan/source/b.md", "rules": []}),
    ]) + "\n")
    findings, scanned = gate.parse_yara(str(ndjson), "/tmp/scan")
    assert scanned == 2
    assert findings == [gate.Finding("yara-x", "R1", "source/a.md")]


def _sarif(results, invocations=None):
    run = {"results": results}
    if invocations is not None:
        run["invocations"] = invocations
    return {"runs": [run]}


def _result(rule, uri):
    return {"ruleId": rule, "message": {"text": "m"},
            "locations": [{"physicalLocation": {
                "artifactLocation": {"uri": uri},
                "region": {"startLine": 3}}}]}


def test_sarif_tool_failures_count_as_findings(tmp_path):
    sarif = tmp_path / "x.sarif"
    sarif.write_text(json.dumps(_sarif([], [{
        "executionSuccessful": False,
        "toolExecutionNotifications": [
            {"level": "error", "descriptor": {"id": "E1"},
             "message": {"text": "boom"}},
            {"level": "warning", "descriptor": {"id": "W1"},
             "message": {"text": "fine"}}]}])))
    rules = [f.rule for f in gate.parse_sarif(str(sarif), "semgrep")]
    assert rules == ["tool-execution-failed", "tool-error:E1"]


def test_filter_sarif_drops_only_accepted_results(tmp_path):
    src, dst = tmp_path / "in.sarif", tmp_path / "out.sarif"
    src.write_text(json.dumps(_sarif([_result("keep", "a.py"),
                                      _result("drop", "b.py")])))
    accepted = [(gate.Finding("semgrep", "drop", "b.py"), {})]
    gate.filter_sarif(str(src), str(dst), accepted)
    kept = json.loads(dst.read_text())["runs"][0]["results"]
    assert [r["ruleId"] for r in kept] == ["keep"]
