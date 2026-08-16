#!/usr/bin/env python3
"""Check every qualified cross-module reference against what that module exports.

    tools/check-refs.py src/*.bas

VBA compiles ONE PROCEDURE AT A TIME, on demand. A reference to a member that
does not exist - `modNote.NOTE_SIZE_AUTO` after the constant has moved to
modOptions - therefore does not fail when the project loads, and does not fail
when the test harness runs every other command. It fails the first time someone
calls the one procedure containing that line, and it fails as

    Compile error in hidden module: modRibbon

which names the module and not the line, and reads like a version or
architecture problem. Exactly that cost a round trip through a real PowerPoint
on 2026-08-16, having passed every test.

So: find `modSomething.Member` everywhere, and require Member to be declared
Public in modSomething.bas. Cheap, and it catches the whole class.
"""
import re
import sys
from pathlib import Path

# Public Sub / Function / Property / Const / Enum / Type, and bare variables.
DECL = re.compile(
    r"^Public\s+(?:Sub|Function|Const|Enum|Type|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)"
    r"|^Public\s+([A-Za-z_]\w*)\s*(?:\(\s*\))?\s+As\s",
    re.MULTILINE,
)
REF = re.compile(r"\b(mod[A-Za-z0-9_]+)\.([A-Za-z_]\w*)")
# A line whose code part is entirely a comment. VBA comments start with '
COMMENT = re.compile(r"^\s*'")


def exports(text):
    found = set()
    for m in DECL.finditer(text):
        found.add(m.group(1) or m.group(2))
    return found


def code_lines(text):
    """Lines with comments and string literals removed, so neither can fake a
    reference. Splitting on quotes and keeping the even parts drops strings."""
    for raw in text.splitlines():
        if COMMENT.match(raw):
            continue
        yield "".join(raw.split('"')[::2])


def main(paths):
    modules = {}
    for p in paths:
        path = Path(p)
        modules[path.stem] = path.read_text(encoding="utf-8", errors="replace")

    api = {name: exports(text) for name, text in modules.items()}

    bad = []
    for name, text in modules.items():
        for lineno, line in enumerate(code_lines(text), 1):
            for target, member in REF.findall(line):
                if target == name or target not in api:
                    # A module not in the set is not this check's business - the
                    # build list check already requires every module to ship.
                    continue
                if member not in api[target]:
                    bad.append(f"{name}.bas: {target}.{member} is not Public in {target}.bas")

    # Line numbers are approximate once comments are dropped, so they are left
    # off rather than printed wrong.
    for msg in sorted(set(bad)):
        print(msg)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
