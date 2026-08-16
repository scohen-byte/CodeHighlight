#!/usr/bin/env python3
"""Reference classification for a source file, as a per-character mask.

`tools/lab.py` is the spec. This prints what it decides, one mask line per
source line, so the VBA scanner can be diffed against it instead of eyeballed:

    python3 tools/lexref.py tests/samples/python/triple.py > expected.txt
    <VBA emits the same format>                             > actual.txt
    diff expected.txt actual.txt

A mask is the same length as its source line, one character per source
character, so a wrong column shows up as a wrong column rather than as a vague
sense that the colours look off.

    .  default        c  comment       s  string        n  number
    k  keyword ctrl   d  keyword decl  f  function      t  class / type
    v  variable      1/2/3  bracket depth

Usage:
    lexref.py FILE...            masks only, machine-comparable
    lexref.py --pretty FILE...   source and mask interleaved, for reading
    lexref.py --legend           the mask alphabet
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lab import classify  # noqa: E402  - same directory, deliberate
from pygments import lex  # noqa: E402
from pygments.lexers import PythonLexer  # noqa: E402

# Palette key -> mask character. One character per class, so masks line up.
KEY2CHAR = {
    "default": ".",
    "comment": "c",
    "string":  "s",
    "number":  "n",
    "kw_ctrl": "k",
    "kw_decl": "d",
    "func":    "f",
    "cls":     "t",
    "var":     "v",
    "br1":     "1",
    "br2":     "2",
    "br3":     "3",
}

# Languages are looked up by id, so a second language means one more entry.
LEXERS = {"python": PythonLexer}


def masks(source: str, language: str = "python") -> list[str]:
    """One mask string per line of source, each as long as its line."""
    lexer_cls = LEXERS[language]
    out: list[str] = [""]
    for key, value in classify(list(lex(source, lexer_cls()))):
        char = KEY2CHAR[key]
        for n, part in enumerate(value.split("\n")):
            if n:
                out.append("")
            # Every character takes its token's class, whitespace included.
            # Tempting to blank spaces for readability, but the renderer colours
            # a whole span - the spaces inside a string ARE string-coloured - so
            # blanking them would make the mask disagree with the thing it is
            # meant to be checking.
            out[-1] += char * len(part)
    # pygments appends a trailing newline, which produces one empty mask.
    if out and not out[-1]:
        out.pop()
    return out


def check_alignment(source: str, mask_lines: list[str]) -> list[str]:
    """Mask and source must agree line for line, character for character."""
    problems = []
    src_lines = source.split("\n")
    if src_lines and src_lines[-1] == "":
        src_lines.pop()
    if len(src_lines) != len(mask_lines):
        problems.append(f"line count: source {len(src_lines)}, mask {len(mask_lines)}")
    for i, (s, m) in enumerate(zip(src_lines, mask_lines), 1):
        if len(s) != len(m):
            problems.append(f"line {i}: source {len(s)} chars, mask {len(m)}")
    return problems


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*", type=Path)
    ap.add_argument("--pretty", action="store_true",
                    help="interleave source and mask")
    ap.add_argument("--legend", action="store_true")
    ap.add_argument("--lang", default="python")
    a = ap.parse_args()

    if a.legend:
        for key, char in KEY2CHAR.items():
            print(f"  {char}  {key}")
        return 0

    status = 0
    for path in a.files:
        source = path.read_text()
        mask_lines = masks(source, a.lang)

        problems = check_alignment(source, mask_lines)
        if problems:
            status = 1
            for p in problems:
                print(f"{path}: MISALIGNED {p}", file=sys.stderr)

        if len(a.files) > 1 or a.pretty:
            print(f"### {path}")
        if a.pretty:
            for n, (src, mask) in enumerate(zip(source.split("\n"), mask_lines), 1):
                print(f"{n:4} | {src}")
                print(f"     | {mask}")
        else:
            for mask in mask_lines:
                print(mask)
    return status


if __name__ == "__main__":
    sys.exit(main())
