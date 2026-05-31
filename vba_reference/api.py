"""High-level query API for the VBA reference library."""

from __future__ import annotations

from typing import Optional

from . import _data
from .models import Constant, Member, MemberRef, TypeDoc


def libraries() -> list[dict]:
    """Return the master catalog: one dict per library with ``folder``,
    ``library``, ``type_count`` and ``types``."""
    return list(_data.master_index()["libraries"])


def library_names() -> list[str]:
    """Return the folder name of every library (e.g. ``"excel"``)."""
    return [lib["folder"] for lib in _data.master_index()["libraries"]]


def list_types(library: str) -> list[dict]:
    """Return ``[{"name", "kind"}, ...]`` for one library."""
    folder = library.lower()
    for lib in _data.master_index()["libraries"]:
        if lib["folder"] == folder:
            return list(lib["types"])
    raise KeyError(f"Unknown library '{library}'.")


def locate_type(name: str) -> list[MemberRef]:
    """Return every place a type ``name`` is defined (across libraries)."""
    return [
        MemberRef(name=canonical, library=folder, type=canonical, kind=kind)
        for folder, canonical, kind in _data.type_locations().get(name.lower(), [])
    ]


def get_type(name: str, library: Optional[str] = None) -> TypeDoc:
    """Return the full :class:`TypeDoc` for ``name``.

    If ``library`` is omitted and the name is defined in more than one library,
    the first match (in catalog order) is returned. Pass ``library`` to
    disambiguate. Lookup is case-insensitive.
    """
    locs = _data.type_locations().get(name.lower())
    if not locs:
        raise KeyError(f"Unknown type '{name}'.")
    if library is not None:
        folder = library.lower()
        match = next((c for f, c, _ in locs if f == folder), None)
        if match is None:
            raise KeyError(f"Type '{name}' is not in library '{library}'.")
        canonical, chosen = match, folder
    else:
        chosen, canonical, _ = locs[0]
    return TypeDoc.from_dict(_data.load_type_json(chosen, canonical))


def get_member(type_name: str, member_name: str,
               library: Optional[str] = None) -> Optional[Member]:
    """Return a single member of a type, or ``None`` if absent."""
    return get_type(type_name, library).member(member_name)


def get_constant(enum_name: str, constant_name: str,
                 library: Optional[str] = None) -> Optional[Constant]:
    """Return a single constant of an enumeration/module, or ``None``."""
    return get_type(enum_name, library).constant(constant_name)


def find_members(name: str, case_sensitive: bool = False) -> list[MemberRef]:
    """Return every definition of a member ``name`` across all libraries.

    Uses the cross-library ``members.json`` index. Matches case-insensitively
    by default.
    """
    if case_sensitive:
        defs = _data.members_index()["members"].get(name)
        exact = name
    else:
        hit = _data.members_lower().get(name.lower())
        if hit is None:
            return []
        exact, defs = hit
    if not defs:
        return []
    return [MemberRef.from_dict(exact, d) for d in defs]


def search_types(substring: str) -> list[MemberRef]:
    """Return type definitions whose name contains ``substring`` (case-insensitive)."""
    low = substring.lower()
    out: list[MemberRef] = []
    for names in _data.type_locations().values():
        for folder, canonical, kind in names:
            if low in canonical.lower():
                out.append(MemberRef(name=canonical, library=folder,
                                     type=canonical, kind=kind))
    out.sort(key=lambda r: (r.name.lower(), r.library))
    return out


def search_members(substring: str) -> list[str]:
    """Return member names containing ``substring`` (case-insensitive), sorted."""
    low = substring.lower()
    names = [exact for exact, _ in _data.members_lower().values()
             if low in exact.lower()]
    names.sort(key=str.lower)
    return names


def data_path() -> str:
    """Return the resolved filesystem location of the reference data."""
    return str(_data.data_root())
