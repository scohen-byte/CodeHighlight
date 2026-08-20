#!/usr/bin/env bash
#
# Thin slice: render one sample as a real code block and bring the PNG back.
#
#   tools/run-slice.sh [sample-name]        default: lab_snippet
#
# Two checks, and the second is the one that matters:
#
#   1. The PNG, for eyeballing against dist/ladder.pptx slide 5 - the look that
#      was signed off.
#   2. A diff of the mask of the SHAPE'S text against the reference mask for the
#      file. Renders cannot be compared by eye closely enough to catch a
#      one-character colour shift, and that shift is the classic failure here:
#      PowerPoint uses CR for paragraph marks and Characters() is 1-based.
#      If this diff is clean, the bridge from scanner to shape is correct.

set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LANG_ID="${LANG_ID:-python}"
NAME="${1:-lab_snippet}"

source "$REPO/tools/win-lab.sh"
WIN_STAGE="$WIN_LAB\\slice"
STAGE="$LAB/slice"

SAMPLES="$REPO/tests/samples/$LANG_ID"
OUT="$REPO/dist/slice"

die() { printf 'run-slice: %s\n' "$*" >&2; exit 1; }

[[ -f "$SAMPLES/$NAME.py" ]] || die "no such sample: $SAMPLES/$NAME.py"

rm -rf "$STAGE"
mkdir -p "$STAGE/src" "$STAGE/out" "$OUT"
cp "$REPO"/src/*.bas "$STAGE/src/"
cp "$SAMPLES/$NAME.py" "$STAGE/"

echo "== PowerPoint =="
powershell.exe -NoProfile -ExecutionPolicy Bypass \
    -File "$(wslpath -w "$REPO/tools/slice-test.ps1")" \
    -SrcDir "$WIN_STAGE\\src" \
    -Sample "$WIN_STAGE\\$NAME.py" \
    -Png    "$WIN_STAGE\\out\\$NAME.png" \
    -Lang   "$LANG_ID" || die "the PowerShell driver failed"

[[ -f "$STAGE/out/$NAME.png" ]] || die "no PNG produced"
cp "$STAGE/out/$NAME.png" "$OUT/"

echo
echo "== shape text vs reference =="
python3 "$REPO/tools/lexref.py" --lang "$LANG_ID" "$SAMPLES/$NAME.py" > "$OUT/$NAME.expected.mask" \
    || die "reference classifier failed"

if [[ -f "$STAGE/out/$NAME.mask" ]]; then
    # The shape mask arrives with CR paragraph marks, since that is what
    # PowerPoint hands back, and with no trailing newline because the block
    # strips trailing blank lines. `sed -e '$a\'` appends a final newline when
    # one is missing, so the two files differ only where the colours differ.
    tr '\r' '\n' < "$STAGE/out/$NAME.mask" | sed -e '$a\' > "$OUT/$NAME.actual.mask"
    sed -i -e '$a\' "$OUT/$NAME.expected.mask"
    if diff -q "$OUT/$NAME.expected.mask" "$OUT/$NAME.actual.mask" >/dev/null; then
        echo "  masks match - no colour shift, the bridge is correct"
    else
        echo "  MASKS DIFFER - the shape text is not what the file said"
        diff "$OUT/$NAME.expected.mask" "$OUT/$NAME.actual.mask" | head -20
    fi
else
    echo "  no shape mask written"
fi

echo
echo "PNG: $OUT/$NAME.png"
