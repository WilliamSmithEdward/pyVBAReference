"""Command-line interface: ``vba-ref`` / ``python -m vba_reference``."""

from __future__ import annotations

import argparse
import sys

from . import api
from .models import Member, TypeDoc


def _print_member(m: Member, indent: str = "") -> None:
    head = m.signature or m.name
    if m.type:
        head = f"{m.name} As {m.type}"
        if m.access:
            head += f"  ({m.access})"
    print(f"{indent}{head}")
    if m.description:
        print(f"{indent}  {m.description}")
    for p in m.parameters:
        opt = "optional" if p.optional else "required"
        t = f" As {p.type}" if p.type else ""
        desc = f": {p.description}" if p.description else ""
        print(f"{indent}  - {p.name}{t} ({opt}){desc}")


def _cmd_libs(_args: argparse.Namespace) -> int:
    for lib in api.libraries():
        print(f"{lib['folder']:<10} {lib['type_count']:>5}  {lib['library']}")
    return 0


def _cmd_where(args: argparse.Namespace) -> int:
    refs = api.find_members(args.name)
    if not refs:
        # Maybe it is a type rather than a member.
        refs = api.locate_type(args.name)
    if not refs:
        print(f"No member or type named '{args.name}' found.", file=sys.stderr)
        return 1
    for r in refs:
        print(f"{r.library:<10} {r.type:<28} {r.kind}")
    return 0


def _cmd_type(args: argparse.Namespace) -> int:
    try:
        t: TypeDoc = api.get_type(args.name, args.library)
    except KeyError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(f"# {t.name}  ({t.kind} - {t.library})")
    if t.description:
        print(t.description)
    if t.constants:
        print(f"\nConstants ({len(t.constants)}):")
        for c in t.constants:
            desc = f" - {c.description}" if c.description else ""
            print(f"  {c.name} = {c.value}{desc}")
    for label, members in (("Properties", t.properties),
                           ("Methods", t.methods),
                           ("Events", t.events),
                           ("Functions", t.functions)):
        if members:
            print(f"\n{label} ({len(members)}):")
            for m in members:
                _print_member(m, indent="  ")
    return 0


def _cmd_member(args: argparse.Namespace) -> int:
    m = api.get_member(args.type, args.member, args.library)
    if m is None:
        print(f"'{args.type}' has no member '{args.member}'.", file=sys.stderr)
        return 1
    _print_member(m)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="vba-ref",
        description="Query the VBA object-model reference.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("libs", help="list libraries").set_defaults(func=_cmd_libs)

    p_where = sub.add_parser(
        "where", help="show where a member or type is defined")
    p_where.add_argument("name")
    p_where.set_defaults(func=_cmd_where)

    p_type = sub.add_parser("type", help="show a full type entry")
    p_type.add_argument("name")
    p_type.add_argument("-l", "--library", default=None)
    p_type.set_defaults(func=_cmd_type)

    p_member = sub.add_parser("member", help="show one member of a type")
    p_member.add_argument("type")
    p_member.add_argument("member")
    p_member.add_argument("-l", "--library", default=None)
    p_member.set_defaults(func=_cmd_member)

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
