#!/usr/bin/env bash
#
# Where the Windows side of these tools works. Sourced, never run:
#
#   source "$(dirname "${BASH_SOURCE[0]}")/win-lab.sh"
#
# Everything here drives PowerPoint over COM, so the files it opens have to sit
# on the Windows filesystem - a \\wsl.localhost path is not somewhere PowerPoint
# can be relied on to read from. That location is per-machine, so it is asked
# for rather than written down. The scripts used to hard-code one developer's
# profile, which built nothing at all on anyone else's machine.
#
# Sets, all absolute:
#
#   WIN_HOME    the Windows profile, Windows-style   C:\Users\sara
#   WIN_LAB     scratch root, Windows-style          C:\Users\sara\ppt-lab
#   LAB         the same scratch root, as WSL sees it
#   WIN_ADDINS  where PowerPoint loads add-ins from, Windows-style
#   ADDINS      the same folder, as WSL sees it
#
# The scratch root is disposable: each script clears its own subdirectory of it
# on every run.

_winlab_die() { printf 'win-lab: %s\n' "$*" >&2; exit 1; }

command -v powershell.exe >/dev/null 2>&1 \
    || _winlab_die "powershell.exe is not on PATH - these tools need WSL on Windows with PowerPoint installed"
command -v wslpath >/dev/null 2>&1 \
    || _winlab_die "wslpath is not on PATH - these tools need to run inside WSL"

# Ask Windows rather than guessing: the profile is not always C:\Users\<name>,
# and the WSL user name need not match the Windows one.
WIN_HOME="$(powershell.exe -NoProfile -NonInteractive -Command 'Write-Output $env:USERPROFILE' 2>/dev/null | tr -d '\r\n')"
[[ -n "$WIN_HOME" ]] || _winlab_die "Windows did not report a USERPROFILE"

WIN_LAB="$WIN_HOME\\ppt-lab"
WIN_ADDINS="$WIN_HOME\\AppData\\Roaming\\Microsoft\\Addins"

LAB="$(wslpath -u "$WIN_LAB")" || _winlab_die "cannot map $WIN_LAB into WSL"
ADDINS="$(wslpath -u "$WIN_ADDINS")" || _winlab_die "cannot map $WIN_ADDINS into WSL"

# wslpath translates a path whether or not it exists, so the profile itself is
# what gets checked. A profile on a drive WSL has not mounted fails here, with
# the reason, rather than fifty lines later as an empty staging directory.
_WINLAB_HOME_WSL="$(wslpath -u "$WIN_HOME")"
[[ -d "$_WINLAB_HOME_WSL" ]] || _winlab_die "the Windows profile $WIN_HOME is not reachable from WSL
     (looked in $_WINLAB_HOME_WSL - is that drive mounted?)"
unset _WINLAB_HOME_WSL
