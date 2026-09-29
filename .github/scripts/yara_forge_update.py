"""Propose the next YARA Forge release for .github/security/yara-forge.toml.

Picks the newest YARA Forge release that is at least COOLDOWN_DAYS old. If
it differs from the pinned one, rewrites the pin with the digest GitHub
records for the release asset (never the hash of the downloaded file),
verifies both the old and the new pack against their digests, and writes a
pull request body listing the rules added and removed.

Writes changed=true|false, tag and previous to $GITHUB_OUTPUT. Needs the gh
CLI, authenticated through GH_TOKEN.

    yara_forge_update.py --pin .github/security/yara-forge.toml --body body.md
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import tomllib
import zipfile
from datetime import datetime, timedelta, timezone

REPO = "YARAHQ/yara-forge"
COOLDOWN_DAYS = 7
DOWNLOAD_PREFIX = f"https://github.com/{REPO}/releases/download/"
RULE_RE = re.compile(rb"^(?:private\s+|global\s+)*rule\s+(\w+)", re.MULTILINE)


def gh_api(path: str):
    out = subprocess.run(["gh", "api", path], check=True, capture_output=True,
                         text=True, encoding="utf-8")
    return json.loads(out.stdout)


def newest_aged_release(now: datetime) -> dict:
    for release in gh_api(f"repos/{REPO}/releases?per_page=30"):
        if release["draft"] or release["prerelease"]:
            continue
        published = datetime.fromisoformat(
            release["published_at"].replace("Z", "+00:00"))
        if now - published >= timedelta(days=COOLDOWN_DAYS):
            return release
    raise SystemExit(f"No {REPO} release is {COOLDOWN_DAYS} days old yet.")


def asset_digest(release: dict, name: str) -> str:
    asset = next((a for a in release["assets"] if a["name"] == name), None)
    if asset is None:
        raise SystemExit(f"Release {release['tag_name']} has no asset {name}.")
    digest = asset.get("digest") or ""
    if not digest.startswith("sha256:"):
        raise SystemExit(f"GitHub records no SHA-256 for {name} in "
                         f"{release['tag_name']}; refusing to pin it.")
    return digest[len("sha256:"):]


def rule_names(tag: str, asset: str, sha256: str) -> set:
    """Download one pack from the release by tag and asset name - no URL is
    taken from the pin file - and return the rule names it defines."""
    data = subprocess.run(
        ["gh", "release", "download", tag, "--repo", REPO,
         "--pattern", asset, "--output", "-"],
        check=True, capture_output=True, timeout=300).stdout
    got = hashlib.sha256(data).hexdigest()
    if got != sha256:
        raise SystemExit(f"{asset} in {tag}: SHA-256 {got} does not match "
                         f"the expected {sha256}.")
    names = set()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for member in zf.namelist():
            if member.endswith(".yar"):
                names.update(m.decode() for m in RULE_RE.findall(zf.read(member)))
    return names


def rewrite_pin(path: str, fields: dict) -> None:
    with open(path, encoding="utf-8") as fh:
        lines = fh.read().splitlines()
    header = []
    for line in lines:
        if line.startswith("#") or not line.strip():
            header.append(line)
        else:
            break
    while header and not header[-1].strip():
        header.pop()
    body = [f'{key} = "{value}"' for key, value in fields.items()]
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(header + [""] + body) + "\n")


def write_output(**values) -> None:
    target = os.environ.get("GITHUB_OUTPUT")
    lines = [f"{key}={value}" for key, value in values.items()]
    if target:
        with open(target, "a", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")
    print("\n".join(lines))


def body_text(old: dict, new: dict, added: set, removed: set) -> str:
    def listing(names: set) -> list:
        shown = sorted(names)[:50]
        out = [f"- `{name}`" for name in shown]
        if len(names) > len(shown):
            out.append(f"- ...and {len(names) - len(shown)} more")
        return out

    lines = [
        f"Moves the YARA Forge `{new['package']}` rule pack from "
        f"`{old['tag']}` to `{new['tag']}` (published {new['published']}, "
        f"the newest release at least {COOLDOWN_DAYS} days old).",
        "",
        f"The new SHA-256 is the digest GitHub records for "
        f"`{new['asset']}`. Both packs were downloaded and checked against "
        f"their digests to compare them.",
        "",
        f"| | `{old['tag']}` | `{new['tag']}` |",
        "| --- | --- | --- |",
        f"| Rules | {new['_old_count']:,} | {new['_new_count']:,} |",
        "",
        f"Release: https://github.com/{REPO}/releases/tag/{new['tag']}",
        "",
    ]
    if added:
        lines += [f"### Added ({len(added):,})", ""] + listing(added) + [""]
    if removed:
        lines += [f"### Removed ({len(removed):,})", ""] + listing(removed) + [""]
    lines += [
        "The Security workflow was started on this branch; its malware scan "
        "compiles the new pack with YARA-X, checks it still detects the EICAR "
        "test file, and scans the repository and the built packages with it. "
        "Merge only when it passes.",
    ]
    return "\n".join(lines) + "\n"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--pin", required=True)
    parser.add_argument("--body", required=True)
    args = parser.parse_args(argv)

    with open(args.pin, "rb") as fh:
        old = tomllib.load(fh)
    release = newest_aged_release(datetime.now(timezone.utc))
    tag = release["tag_name"]
    if tag == old["tag"]:
        print(f"YARA Forge {tag} is already pinned.")
        write_output(changed="false", tag=tag, previous=old["tag"])
        return 0

    published = release["published_at"][:10]
    if published < old["published"]:
        print(f"Newest aged release {tag} is older than the pinned "
              f"{old['tag']}; leaving the pin alone.")
        write_output(changed="false", tag=old["tag"], previous=old["tag"])
        return 0

    new = {
        "tag": tag,
        "published": published,
        "package": old["package"],
        "asset": old["asset"],
        "url": f"{DOWNLOAD_PREFIX}{tag}/{old['asset']}",
        "sha256": asset_digest(release, old["asset"]),
    }
    old_rules = rule_names(old["tag"], old["asset"], old["sha256"])
    new_rules = rule_names(new["tag"], new["asset"], new["sha256"])
    rewrite_pin(args.pin, new)

    body = body_text(old, dict(new, _old_count=len(old_rules),
                               _new_count=len(new_rules)),
                     new_rules - old_rules, old_rules - new_rules)
    with open(args.body, "w", encoding="utf-8") as fh:
        fh.write(body)
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as fh:
            fh.write(body)
    write_output(changed="true", tag=tag, previous=old["tag"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
