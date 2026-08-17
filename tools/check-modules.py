#!/usr/bin/env python3
"""Check that every harness imports every module its imports actually reach.

    tools/check-modules.py

A PowerShell harness lists the .bas files it imports into a scratch deck. If a
listed module references one that is NOT listed, VBA has no module of that name
and reports the reference as "Variable not defined" - naming a local variable
that is nothing of the sort. That is a confusing enough error on its own; what
makes it worth a check is that it appears only when a module gains a NEW
cross-module call, so it turns up long after the harness list was written and
looks like a bug in the code being tested.

tools/check-refs.py resolves references against src/. This resolves them
against each harness's own list, which is a different question with the same
failure mode - the build's module list had exactly this bug once already.
"""
import re
import sys
from pathlib import Path

REF = re.compile(r"\b(mod[A-Za-z0-9_]+)\.")
COMMENT = re.compile(r"^\s*'")


def references(text, own, known):
    """Modules this one calls into. Comments and string literals dropped, so
    neither can invent a dependency."""
    out = set()
    for raw in text.splitlines():
        if COMMENT.match(raw):
            continue
        code = "".join(raw.split('"')[::2])
        for name in REF.findall(code):
            if name in known and name != own:
                out.add(name)
    return out


def main():
    root = Path(__file__).resolve().parent.parent
    mods = {p.stem: p.read_text(encoding="utf-8", errors="replace")
            for p in (root / "src").glob("*.bas")}
    graph = {n: references(t, n, mods) for n, t in mods.items()}

    bad = False
    for ps1 in sorted((root / "tools").glob("*.ps1")):
        m = re.search(r"\$MODULES\s*=\s*@\((.*?)\)", ps1.read_text(), re.S)
        if not m:
            continue
        listed = set(re.findall(r"'(mod[A-Za-z0-9_]+)'", m.group(1)))

        # Everything the listed set can reach, transitively.
        closure, todo = set(listed), list(listed)
        while todo:
            for dep in graph.get(todo.pop(), ()):
                if dep not in closure:
                    closure.add(dep)
                    todo.append(dep)

        missing = closure - listed
        if missing:
            print("%s: $MODULES is missing %s, which its listed modules call into"
                  % (ps1.name, ", ".join(sorted(missing))))
            bad = True
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
