"""vba_reference - programmatic access to the VBA object-model reference.

Quick start::

    import vba_reference as vba

    vba.library_names()                 # ['excel', 'office', ...]
    ws = vba.get_type("Worksheet")      # TypeDoc
    ws.member("Protect").parameters     # tuple[Parameter, ...]
    vba.find_members("MsgBox")          # [MemberRef(library='vba', type='Interaction', ...)]
    vba.get_constant("XlFileFormat", "xlCSV").value  # 6
"""

from __future__ import annotations

from .api import (
    data_path,
    find_members,
    get_constant,
    get_member,
    get_type,
    libraries,
    library_names,
    list_types,
    locate_type,
    search_members,
    search_types,
)
from .models import Constant, Member, MemberRef, Parameter, TypeDoc

__version__ = "1.0.0"

__all__ = [
    "__version__",
    # models
    "Parameter",
    "Member",
    "Constant",
    "TypeDoc",
    "MemberRef",
    # api
    "libraries",
    "library_names",
    "list_types",
    "locate_type",
    "get_type",
    "get_member",
    "get_constant",
    "find_members",
    "search_types",
    "search_members",
    "data_path",
]
