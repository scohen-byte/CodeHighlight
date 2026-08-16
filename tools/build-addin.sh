#!/usr/bin/env bash
#
# Build CodeHighlight.ppam from source, ribbon and all.
#
#   tools/build-addin.sh              build only, into dist/
#   tools/build-addin.sh --install    build, then replace the installed add-in
#
# The build is from scratch every time: a new presentation, the shipping modules
# imported into it, saved as a .ppam, then the ribbon injected. Nothing mutates
# the installed copy until --install, so a bad build cannot break a good install.
#
# PowerPoint must be CLOSED. It holds both the add-in and any open deck, and a
# repack behind its back gets silently overwritten.

set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
NAME="CodeHighlight.ppam"

WIN_SRC='C:\Users\User\ppt-lab\build\src'
WIN_OUT='C:\Users\User\ppt-lab\build\CodeHighlight.pptm'
STAGE="/mnt/c/Users/User/ppt-lab/build"
INSTALL_DIR="/mnt/c/Users/User/AppData/Roaming/Microsoft/Addins"

die() { printf 'build-addin: %s\n' "$*" >&2; exit 1; }

INSTALL=0
[[ "${1:-}" == "--install" ]] && INSTALL=1

# pgrep cannot see Windows processes from WSL, so ask Windows. Exit 1 means
# PowerPoint is running.
powershell.exe -NoProfile -Command \
    'if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }' \
    >/dev/null 2>&1 \
    || die "PowerPoint is running. Close it first - it holds the add-in open."

rm -rf "$STAGE"
mkdir -p "$STAGE/src"
cp "$REPO"/src/*.bas "$STAGE/src/"

echo "== build =="
powershell.exe -NoProfile -ExecutionPolicy Bypass \
    -File "$(wslpath -w "$REPO/tools/build-addin.ps1")" \
    -SrcDir "$WIN_SRC" -Out "$WIN_OUT" || die "the build failed"

[[ -f "$STAGE/CodeHighlight.pptm" ]] || die "no .pptm produced"

echo
echo "== convert to add-in =="
# A .ppam is a .pptm whose main part carries the add-in content type. That one
# string is the whole difference, and PowerPoint will not write it over COM.
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
unzip -qq "$STAGE/CodeHighlight.pptm" -d "$WORK/pkg"
CT="$WORK/pkg/[Content_Types].xml"
sed -i 's#application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml#application/vnd.ms-powerpoint.addin.macroEnabled.main+xml#' "$CT"
grep -q 'addin.macroEnabled.main' "$CT" || die "content type not converted"

# Same packing rules as pack-ribbon.sh: [Content_Types].xml first, no directory
# entries, and -nw because the brackets are zip wildcard syntax.
(
    cd "$WORK/pkg"
    zip -q -X -D -nw "$WORK/out.ppam" '[Content_Types].xml'
    zip -q -r -X -D -nw "$WORK/out.ppam" . -x '[Content_Types].xml'
)
cp -f "$WORK/out.ppam" "$STAGE/$NAME"
echo "  content type -> addin.macroEnabled.main"

[[ -f "$STAGE/$NAME" ]] || die "no .ppam produced"

echo
echo "== ribbon =="
"$REPO/tools/pack-ribbon.sh" "$STAGE/$NAME" || die "packing the ribbon failed"

mkdir -p "$REPO/dist"
cp "$STAGE/$NAME" "$REPO/dist/$NAME"
echo
echo "built: $REPO/dist/$NAME"

if [[ $INSTALL -eq 1 ]]; then
    [[ -d "$INSTALL_DIR" ]] || die "no add-ins folder at $INSTALL_DIR"
    if [[ -f "$INSTALL_DIR/$NAME" ]]; then
        cp -f "$INSTALL_DIR/$NAME" "$INSTALL_DIR/$NAME.prev"
        echo "previous install kept as $NAME.prev"
    fi
    cp -f "$REPO/dist/$NAME" "$INSTALL_DIR/$NAME"
    echo "installed: $INSTALL_DIR/$NAME"
    echo "Reopen PowerPoint. If the tab misbehaves, untick and re-tick the add-in."
fi
