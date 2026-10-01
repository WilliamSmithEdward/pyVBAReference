# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/pyVBAReference/security/advisories/new).
Only the maintainer sees it. Include the `vba-reference` version, the
Python version, the call or `vba-ref` command that shows the problem, and
the smallest steps that reproduce it, with credentials and private data
removed.

A confirmed vulnerability is fixed in a release on PyPI, and the advisory
is published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on PyPI receives security fixes. Older releases are
not maintained separately; update when a fix ships.

## Scope

The installed package, `vba-reference`, reads JSON files bundled inside it.
It has no runtime dependencies, opens no network connections, and executes
nothing from its data. A way to make it read beyond its bundled data, reach
the network or run code is a vulnerability.

The generator scripts in the repository root (`scrape_excel_object_model.py`,
`mslearn_docs.py`, `consolidate_reference.py`) are not part of the installed
package. They run by hand on Windows, read the registered COM type
libraries, and download the MicrosoftDocs/VBA-Docs archive from GitHub over
HTTPS. A flaw in them that a tampered download could exploit is in scope.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily, and again from the
Publish workflow for every release.

- **Code:** CodeQL with the `security-extended` query suite, for Python and
  the GitHub Actions workflows, and Semgrep with the `p/python`,
  `p/security-audit`, `p/secrets` and `p/github-actions` rule sets. Results
  go to the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** there is no dependency audit. The installed package has
  no runtime dependencies (`dependencies = []` in `pyproject.toml`), and the
  tools the workflows install are hash-locked and moved by Dependabot.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan the committed files and the wheel and sdist built from
  them, both as archives and unpacked. ClamAV runs with PUA detection and
  alerts for broken, encrypted, macro-bearing and limit-exceeding files on.
  YARA-X runs the full YARA Forge pack. Each scanner must first detect the
  EICAR test file, and a scanner that errors or cannot update its signatures
  fails the job, so a scanner that detects nothing cannot pass as a clean
  result.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Some of its checks do not fit this project: a single maintainer cannot
  have a second person approve every change, and the package parses no
  input it does not ship (its lookups key dictionaries built from its own
  index), so it is not fuzzed.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[.github/security/known-findings.toml](.github/security/known-findings.toml).
An entry matches the tool, the rule and a glob over the reported path, and
an entry that no longer matches fails the report. zizmor keeps its
exceptions in `.github/zizmor.yml` or inline beside the line they excuse,
each with its reason. The current entries:

| Tool | Rule | Where | Why it is acceptable |
| ---- | ---- | ----- | -------------------- |
| Semgrep | `dynamic-urllib-use-detected` | `mslearn_docs.py` | The only URL passed is a hard-coded HTTPS constant, and `_download()` refuses any non-HTTPS URL. The generator is not in the installed package. |
| YARA-X | `SIGNATURE_BASE_Powershell_Case_Anomaly` | `reference/agentic_llm_primer.md` | A YARA-X 1.20.0 false positive: the rule does not match this file when compiled alone, only when compiled with the rest of the pack. |
| zizmor | `self-repository` | [.github/zizmor.yml](.github/zizmor.yml) | Turned off until GitHub's documentation confirms the `$/` self-repository syntax for reusable workflows called from `publish.yml`. |

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, and the YARA-X engine and YARA Forge rules to a
release and its SHA-256. ClamAV's signatures change too often to pin, so
freshclam fetches and verifies them on every run. The Semgrep rule sets are
fetched from the Semgrep registry on every run, and the security report
records the engine version and which rule sets ran.

Dependabot proposes updates to the GitHub Actions, the scanner images and
the Python tool locks in `.github/requirements` once a version is a week
old, and at once for a security advisory. The Update YARA rules workflow
proposes new YARA pins each week. A minor or patch update, and the YARA
pull request, merges itself once CI, Security and Malware scan pass; a
third-party major version waits for review.

## Releases

Publishing a GitHub release starts the Publish workflow. It runs the tests,
checks that the release tag matches the version in `pyproject.toml`, builds
the wheel and sdist, checks their metadata, and exercises the installed
wheel outside the source tree. It runs Security and Malware scan on the
release commit, and uploads to PyPI through trusted publishing only when
the build and both scans pass. Started by hand, it is a dry run that
publishes nothing.

Every release after v1.0.0 carries `security-report-vX.Y.Z.md` and
`malware-report-vX.Y.Z.md`: the scans of that release's commit, with the
tool, rule and signature versions they used and every finding they
accepted. The reports are attached whether the scans passed or failed.

### Verifying a download

Every file on PyPI carries PyPI's own provenance, which names this
repository's `publish.yml` as the publisher; the file's page on PyPI shows it.
Releases published after 2026-09-30 also carry a GitHub build provenance
attestation, which you can check against any copy of the file:

```
pip download vba-reference --no-deps -d check
gh attestation verify check/<file> --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
