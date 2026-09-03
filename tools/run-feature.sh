#!/usr/bin/env bash
#
# Drive one modSelfTest function against a real PowerPoint and bring the PNG back.
#
#   tools/run-feature.sh NoteTest [sample-name]     default sample: lab_snippet
#
# Run tools/check-vba.sh first. A VBA compile error does not come back as an
# error here - it opens the VBA editor with a modal dialog, and every later COM
# call blocks, which from the outside is indistinguishable from a hang.

set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LANG_ID="${LANG_ID:-python}"
PROC="${1:?usage: run-feature.sh <SelfTestFunction> [sample]}"
NAME="${2:-lab_snippet}"

source "$REPO/tools/win-lab.sh"

# PowerPoint's COM server is single-instance, so a stranded windowless one from
# an earlier run is not just wasted memory - New-Object hands this run that same
# process. Sweep before and after; only windowless ones are touched, so a deck
# someone has open is never at risk.
"$REPO/tools/kill-orphans.sh" --quiet
trap '"$REPO/tools/kill-orphans.sh" --quiet' EXIT
WIN_STAGE="$WIN_LAB\\feature"
STAGE="$LAB/feature"

SAMPLES="$REPO/tests/samples/$LANG_ID"
OUT="$REPO/dist/slice"
PNG_NAME="$(echo "$PROC" | tr '[:upper:]' '[:lower:]').png"

die() { printf 'run-feature: %s\n' "$*" >&2; exit 1; }

[[ -f "$SAMPLES/$NAME.py" ]] || die "no such sample: $SAMPLES/$NAME.py"

rm -rf "$STAGE"
mkdir -p "$STAGE/src" "$STAGE/out" "$OUT"
cp "$REPO"/src/*.bas "$STAGE/src/"
cp "$SAMPLES/$NAME.py" "$STAGE/"

powershell.exe -NoProfile -ExecutionPolicy Bypass \
    -File "$(wslpath -w "$REPO/tools/feature-test.ps1")" \
    -SrcDir "$WIN_STAGE\\src" \
    -Proc   "$PROC" \
    -Sample "$WIN_STAGE\\$NAME.py" \
    -Png    "$WIN_STAGE\\out\\$PNG_NAME" || die "the PowerShell driver failed"

if [[ -f "$STAGE/out/$PNG_NAME" ]]; then
    cp "$STAGE/out/$PNG_NAME" "$OUT/"
    echo "png: $OUT/$PNG_NAME"
fi
