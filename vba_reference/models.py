"""Typed data models for the VBA reference library.

Each model mirrors the JSON shape produced by ``scrape_excel_object_model.py``
and exposes a ``from_dict`` constructor.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional


@dataclass(frozen=True)
class Parameter:
    """A single parameter of a method, function, or event."""

    name: str
    type: str = ""
    optional: bool = False
    description: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> "Parameter":
        return cls(
            name=d.get("name", ""),
            type=d.get("type", ""),
            optional=bool(d.get("optional", False)),
            description=d.get("description", ""),
        )


@dataclass(frozen=True)
class Member:
    """A property, method, event, or module-level function."""

    name: str
    kind: str = ""
    description: str = ""
    signature: str = ""
    type: str = ""          # property data type
    access: str = ""        # property access: read-only / write-only / read/write
    returns: str = ""       # method/function return type ("void" for Sub)
    parameters: tuple[Parameter, ...] = ()

    @classmethod
    def from_dict(cls, d: dict) -> "Member":
        return cls(
            name=d.get("name", ""),
            kind=d.get("kind", ""),
            description=d.get("description", ""),
            signature=d.get("signature", ""),
            type=d.get("type", ""),
            access=d.get("access", ""),
            returns=d.get("returns", ""),
            parameters=tuple(
                Parameter.from_dict(p) for p in d.get("parameters", [])
            ),
        )


@dataclass(frozen=True)
class Constant:
    """An enumeration or module constant."""

    name: str
    value: Any = None
    description: str = ""
    type: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> "Constant":
        return cls(
            name=d.get("name", ""),
            value=d.get("value"),
            description=d.get("description", ""),
            type=d.get("type", ""),
        )


@dataclass(frozen=True)
class TypeDoc:
    """A complete reference entry for one VBA type (class, interface,
    enumeration, or module)."""

    name: str
    kind: str
    library: str
    description: str = ""
    remarks: str = ""
    example: str = ""
    guid: str = ""
    properties: tuple[Member, ...] = ()
    methods: tuple[Member, ...] = ()
    events: tuple[Member, ...] = ()
    functions: tuple[Member, ...] = ()
    constants: tuple[Constant, ...] = ()

    @classmethod
    def from_dict(cls, d: dict, library: Optional[str] = None) -> "TypeDoc":
        return cls(
            name=d.get("name", ""),
            kind=d.get("kind", ""),
            library=library or d.get("library", ""),
            description=d.get("description", ""),
            remarks=d.get("remarks", ""),
            example=d.get("example", ""),
            guid=d.get("guid", ""),
            properties=tuple(Member.from_dict(m) for m in d.get("properties", [])),
            methods=tuple(Member.from_dict(m) for m in d.get("methods", [])),
            events=tuple(Member.from_dict(m) for m in d.get("events", [])),
            functions=tuple(Member.from_dict(m) for m in d.get("functions", [])),
            constants=tuple(Constant.from_dict(c) for c in d.get("constants", [])),
        )

    def all_members(self) -> tuple[Member, ...]:
        """Every callable/queryable member across properties, methods, events,
        and module functions."""
        return self.properties + self.methods + self.events + self.functions

    def member(self, name: str) -> Optional[Member]:
        """Return the member with ``name`` (case-insensitive), or ``None``."""
        low = name.lower()
        for m in self.all_members():
            if m.name.lower() == low:
                return m
        return None

    def constant(self, name: str) -> Optional[Constant]:
        """Return the constant with ``name`` (case-insensitive), or ``None``."""
        low = name.lower()
        for c in self.constants:
            if c.name.lower() == low:
                return c
        return None


@dataclass(frozen=True)
class MemberRef:
    """A pointer to where a member name is defined."""

    name: str
    library: str
    type: str
    kind: str

    @classmethod
    def from_dict(cls, name: str, d: dict) -> "MemberRef":
        return cls(
            name=name,
            library=d.get("library", ""),
            type=d.get("type", ""),
            kind=d.get("kind", ""),
        )
