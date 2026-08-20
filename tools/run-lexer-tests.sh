#!/usr/bin/env bash
#
# Diff the VBA scanner against the reference classifier, one sample at a time.
#
#   tools/run-lexer-tests.sh [sample-name ...]
#
# With no arguments it runs the whole corpus. With names it runs just those,
# which is what you want while chasing one bug.
#
# The loop is: stage the .bas files and the samples onto the C: drive, drive
# PowerPoint over COM to import and run them, copy the masks back, diff against
# tools/lexref.py. A failing sample prints the first differing lines with the
# source line above them, because "line 41 column 12" is not much use on its own.

set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LANG_ID="${LANG_ID:-python}"

source "$REPO/tools/win-lab.sh"
WIN_STAGE="$WIN_LAB\\lexer"
STAGE="$LAB/lexer"

SAMPLES="$REPO/tests/samples/$LANG_ID"
EXPECTED="$REPO/dist/lexer/expected"
ACTUAL="$REPO/dist/lexer/actual"

die() { printf 'run-lexer-tests: %s\n' "$*" >&2; exit 1; }

[[ -d "$SAMPLES" ]] || die "no samples for language '$LANG_ID' at $SAMPLES"

# Which samples to run.
if [[ $# -gt 0 ]]; then
    NAMES=("$@")
else
    NAMES=()
    for f in "$SAMPLES"/*.py; do NAMES+=("$(basename "$f" .py)"); done
fi

rm -rf "$STAGE" "$EXPECTED" "$ACTUAL"
mkdir -p "$STAGE/src" "$STAGE/samples" "$STAGE/masks" "$EXPECTED" "$ACTUAL"

cp "$REPO"/src/*.bas "$STAGE/src/"
for n in "${NAMES[@]}"; do
    [[ -f "$SAMPLES/$n.py" ]] || die "no such sample: $n"
    cp "$SAMPLES/$n.py" "$STAGE/samples/"
done

echo "== VBA =="
powershell.exe -NoProfile -ExecutionPolicy Bypass \
    -File "$(wslpath -w "$REPO/tools/lexer-test.ps1")" \
    -SrcDir    "$WIN_STAGE\\src" \
    -SampleDir "$WIN_STAGE\\samples" \
    -OutDir    "$WIN_STAGE\\masks" \
    -Lang      "$LANG_ID" || die "the PowerShell driver failed"

echo
echo "== diff vs tools/lexref.py =="
fail=0
for n in "${NAMES[@]}"; do
    python3 "$REPO/tools/lexref.py" --lang "$LANG_ID" "$SAMPLES/$n.py" > "$EXPECTED/$n.mask" \
        || die "reference classifier failed on $n"
    # The VBA writes ANSI with LF; strip any stray CR before comparing.
    tr -d '\r' < "$STAGE/masks/$n.mask" > "$ACTUAL/$n.mask" 2>/dev/null \
        || { echo "  $n: NO OUTPUT from the VBA side"; fail=1; continue; }

    if diff -q "$EXPECTED/$n.mask" "$ACTUAL/$n.mask" >/dev/null; then
        printf '  %-14s ok\n' "$n"
    else
        fail=1
        printf '  %-14s DIFFERS\n' "$n"
        # Show the first few differing lines with their source, which is the
        # only form of this output anyone can actually act on.
        awk -v src="$SAMPLES/$n.py" -v want="$EXPECTED/$n.mask" -v got="$ACTUAL/$n.mask" '
            BEGIN {
                while ((getline line < src) > 0) s[++ns] = line
                while ((getline line < want) > 0) e[++ne] = line
                while ((getline line < got) > 0) a[++na] = line
                shown = 0
                n = (ne > na ? ne : na)
                for (i = 1; i <= n && shown < 4; i++) {
                    if (e[i] != a[i]) {
                        shown++
                        printf "    line %d: %s\n", i, s[i]
                        printf "      want   %s\n", e[i]
                        printf "      got    %s\n", a[i]
                    }
                }
                if (ne != na) printf "    line count: want %d, got %d\n", ne, na
            }'
    fi
done

echo
if [[ $fail -eq 0 ]]; then
    echo "all samples match the reference"
else
    echo "masks differ. expected: $EXPECTED  actual: $ACTUAL"
fi
exit $fail
