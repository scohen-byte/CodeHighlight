#!/usr/bin/env bash
#
# Kill PowerPoint processes the test harness left behind.
#
#   tools/kill-orphans.sh            report and kill orphans
#   tools/kill-orphans.sh --quiet    same, but silent when there are none
#
# Every harness run drives a fresh PowerPoint over COM. When one exits badly -
# a VBA error, a timeout, a killed script - the COM server survives with no
# window, holding its memory. They accumulate: a long session can leave a dozen
# invisible PowerPoints resident, which is a lot of RAM for nothing, and there
# is nothing on screen to close.
#
# ONLY WINDOWLESS ONES. A PowerPoint with a window title is a deck someone has
# open, and killing that loses unsaved work. build-addin.sh draws the same
# distinction when it reports what is holding the add-in.

set -uo pipefail

QUIET=0
[[ "${1:-}" == "--quiet" ]] && QUIET=1

command -v powershell.exe >/dev/null 2>&1 || exit 0

OUT=$(powershell.exe -NoProfile -NonInteractive -Command '
    $orphans = @(Get-Process POWERPNT -ErrorAction SilentlyContinue |
                 Where-Object { -not $_.MainWindowTitle })
    if ($orphans.Count -eq 0) { "none"; exit 0 }
    $mb = [math]::Round((($orphans | Measure-Object WorkingSet64 -Sum).Sum / 1MB))
    $n  = $orphans.Count
    $orphans | Stop-Process -Force -ErrorAction SilentlyContinue
    "killed $n orphan(s), freeing about $mb MB"
' 2>/dev/null | tr -d '\r')

case "$OUT" in
    none|"") [[ $QUIET -eq 1 ]] || echo "kill-orphans: none found" ;;
    *)       echo "kill-orphans: $OUT" ;;
esac
