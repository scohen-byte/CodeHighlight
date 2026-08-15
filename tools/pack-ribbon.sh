#!/usr/bin/env bash
#
# Inject the custom ribbon into a .ppam add-in.
#
# A .ppam is an ordinary zip. PowerPoint's UI offers no way to add a ribbon tab
# to one, so the customUI part and its relationship are written in by hand here.
#
#   tools/pack-ribbon.sh /mnt/c/Users/User/AppData/Roaming/Microsoft/Addins/CodeHighlight.ppam
#
# Safe to re-run: any previous ribbon part and relationship are replaced, not
# duplicated. The previous file is kept alongside as <name>.ppam.bak.
#
# PowerPoint caches an add-in while it is loaded, so it must be CLOSED before
# repacking. Repacking a file PowerPoint has open silently gets overwritten back.

set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RIBBON_DEFAULT="$REPO/ribbon/customUI14.xml"

REL_ID="rIdCodeHighlightUI"
REL_TYPE="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
REL_TARGET="customUI/customUI14.xml"

die() { printf 'pack-ribbon: %s\n' "$*" >&2; exit 1; }
note() { printf '  %s\n' "$*"; }

TARGET="${1:-}"
RIBBON="${2:-$RIBBON_DEFAULT}"

[[ -n "$TARGET" ]] || die "usage: pack-ribbon.sh <path-to.ppam> [customUI14.xml]"
[[ -f "$TARGET" ]] || die "no such file: $TARGET"
[[ -f "$RIBBON" ]] || die "no ribbon xml: $RIBBON"
unzip -tqq "$TARGET" >/dev/null 2>&1 || die "not a valid zip (is it really a .ppam?): $TARGET"

# Fail early and loudly rather than producing an add-in with no macros in it.
if ! unzip -l "$TARGET" | grep -q 'vbaProject.bin'; then
    die "no vbaProject.bin inside $TARGET
     The .ppam must be created and saved from PowerPoint on Windows first, with
     at least one module in it. The VBA project binary cannot be built from WSL."
fi

VBA_BEFORE=$(unzip -l "$TARGET" | awk '/vbaProject.bin/ {print $1}')

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

echo "packing $(basename "$TARGET")"
unzip -qq "$TARGET" -d "$WORK/pkg"

# --- the ribbon part ----------------------------------------------------------
mkdir -p "$WORK/pkg/customUI"
cp "$RIBBON" "$WORK/pkg/customUI/customUI14.xml"
note "customUI/customUI14.xml  <- $(basename "$RIBBON")"

# --- the relationship ---------------------------------------------------------
# Without this the part is inert: PowerPoint only looks for a ribbon it is
# pointed at from the package root .rels.
RELS="$WORK/pkg/_rels/.rels"
[[ -f "$RELS" ]] || die "package has no _rels/.rels - corrupt .ppam"

# Drop any existing extensibility relationship, ours or a previous tool's, so
# re-running cannot accumulate duplicates (which makes Office ignore all of them).
sed -i 's#<Relationship[^>]*ui/extensibility[^>]*/>##g' "$RELS"
sed -i "s#</Relationships>#<Relationship Id=\"$REL_ID\" Type=\"$REL_TYPE\" Target=\"$REL_TARGET\"/></Relationships>#" "$RELS"
grep -q "$REL_ID" "$RELS" || die "failed to write the relationship into _rels/.rels"
note "_rels/.rels              <- $REL_ID"

# --- content type -------------------------------------------------------------
# The part is served by the default xml content type. Present in every stock
# package, but assert it rather than assume it.
CT="$WORK/pkg/[Content_Types].xml"
if ! grep -q 'Extension="xml"' "$CT"; then
    # Defaults must precede Overrides, so insert immediately after <Types ...>.
    sed -i 's#\(<Types[^>]*>\)#\1<Default Extension="xml" ContentType="application/xml"/>#' "$CT"
    note "[Content_Types].xml      <- Default xml"
fi

# --- repack -------------------------------------------------------------------
cp -f "$TARGET" "$TARGET.bak"
OUT="$WORK/out.ppam"
( cd "$WORK/pkg" && zip -q -r -X "$OUT" . )
cp -f "$OUT" "$TARGET"

# --- verify -------------------------------------------------------------------
# Cheap, and it catches the one failure that matters: a repack that dropped the
# VBA project would produce an add-in that loads a ribbon wired to nothing.
unzip -tqq "$TARGET" >/dev/null || die "repacked file is not a valid zip"
unzip -l "$TARGET" | grep -q 'customUI/customUI14.xml' || die "ribbon part missing after repack"
VBA_AFTER=$(unzip -l "$TARGET" | awk '/vbaProject.bin/ {print $1}')
[[ "$VBA_BEFORE" == "$VBA_AFTER" ]] || die "vbaProject.bin changed size ($VBA_BEFORE -> $VBA_AFTER) - aborting, restore from $TARGET.bak"

echo "ok. backup at $(basename "$TARGET").bak"
echo "Reload the add-in in PowerPoint (untick and re-tick it under Add-ins) to see the change."
