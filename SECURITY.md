# Security policy

## Reporting a vulnerability

Report it privately through GitHub:
[report a vulnerability](https://github.com/WilliamSmithEdward/pyVBAReference/security/advisories/new).
The report is visible only to the maintainer until an advisory is published.
Please do not open a public issue for a security problem.

## Supported versions

| Version | Supported |
| ------- | --------- |
| 1.x     | yes       |

## What the package does

The installed package, `vba-reference`, reads JSON files bundled inside it.
It has no runtime dependencies, opens no network connections, and executes
nothing from its data.

The generator scripts in the repository root are not part of the installed
package. They run by hand on Windows, read the registered COM type libraries,
and download the MicrosoftDocs/VBA-Docs archive from GitHub over HTTPS.

## How the repository is scanned

The [Security workflow](.github/workflows/security.yml) (CodeQL, Semgrep) and
the [Malware scan workflow](.github/workflows/malware-scan.yml) (ClamAV,
YARA-X) run on every push to `main`, on every pull request, daily, and for
every release.

| Check | What it covers |
| ----- | -------------- |
| CodeQL | Python and the GitHub Actions workflows, with the `security-extended` query suite |
| Semgrep | The `p/python`, `p/security-audit`, `p/secrets` and `p/github-actions` rule sets |
| ClamAV | The committed files and the built wheel and sdist, with signatures updated at the start of every run; PUA detection and alerts for broken, encrypted, macro-bearing and limit-exceeding files are on |
| YARA-X | The same files, unpacked, against the full [YARA Forge](https://github.com/YARAHQ/yara-forge) pack - about 12,000 rules collected from public rule repositories |

The build fails on:

- any finding not accepted in
  [known-findings.toml](.github/security/known-findings.toml)
- an accepted entry that no longer matches anything, so the list cannot
  outlive its reasons
- a scanner that errors, cannot update its signatures, or fails to detect the
  EICAR test file each run starts with - a scanner that detects nothing would
  otherwise look like a clean result

Publishing to PyPI waits for both workflows, so a release with an unexpected
finding is not published.

## Security reports on releases

Every release after v1.0.0 has a `security-report-vX.Y.Z.md` attached, and
later releases a `malware-report-vX.Y.Z.md` beside it: the scans of that
release's commit, with the tool, rule and signature versions it
used and every finding it accepted. The reports are attached whether the scans
passed or failed.

## Known acceptable findings

These are reviewed and accepted. The full reasons are in
[known-findings.toml](.github/security/known-findings.toml), which the build
reads.

| Tool | Rule | Where | Why it is acceptable |
| ---- | ---- | ----- | -------------------- |
| Semgrep | `dynamic-urllib-use-detected` | `mslearn_docs.py` | The only URL passed is a hard-coded HTTPS constant, and `_download()` refuses any non-HTTPS URL. The generator is not in the installed package. |
| YARA-X | `SIGNATURE_BASE_Powershell_Case_Anomaly` | `reference/agentic_llm_primer.md` | A YARA-X 1.20.0 false positive: the rule does not match this file when compiled alone, only when compiled with the rest of the pack. |

## Pinning and updates

Everything a workflow runs is pinned: actions by commit SHA, container images
by digest, Python tools by exact version and hash, the YARA-X binary by
version and SHA-256, and the YARA Forge pack by release and SHA-256.

[Dependabot](.github/dependabot.yml) proposes updates to the actions, images
and Python tools, and a [weekly workflow](.github/workflows/update-yara-rules.yml)
proposes the next YARA Forge release. Nothing is proposed until it is seven
days old, and every proposal is scanned by the Security or Malware scan workflow before it
can be merged.

ClamAV signatures and the Semgrep rule sets change too often to pin, so they
are fetched fresh on every run. freshclam verifies each database's signature,
and the report records the database version used. For Semgrep it records the
engine version and which rule sets ran.

## Repository settings

- `main` is protected by a ruleset: changes arrive through pull requests, the
  **Security passed** and **Malware scan passed** checks must pass before
  merging, and the branch cannot be force-pushed or deleted. Each gate job
  depends on every job in its workflow, so they are the only checks the
  ruleset names, and renaming a scan job never loosens the
  protection. The check is tied to GitHub Actions, so nothing else can report
  it. Repository admins can bypass the ruleset.
- GitHub Actions refuses any action not pinned to a full commit SHA, so a
  tag or branch reference fails the run instead of relying on review to catch
  it.
- Private vulnerability reporting, secret scanning with push protection, and
  Dependabot alerts and security updates are enabled.
