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

# Every script here derives its staging directory from $LAB and clears that
# directory recursively on each run, so $LAB is the one value that must not be
# allowed to come out wrong. On 2026-09-03 a hand-rolled version of this
# translation - `echo` on a backslash path, under zsh - produced a string
# holding a NUL byte. The shell showed a full-length path; execve truncated the
# argument at the NUL and the delete ran against /mnt/c instead. Nothing in the
# printed value gave that away, so the check is made here, once, rather than
# trusted at each call site.
_winlab_check() {
    local val="$1" what="$2"

    # What an exec'd program actually receives. A NUL, or anything else that
    # does not survive the trip through argv, shows up as a shorter string.
    [[ "$(/usr/bin/env printf '%s' "$val")" == "$val" ]] \
        || _winlab_die "$what does not survive exec intact - it holds an embedded NUL or similar.
     Printing it will look right; it will not be what a command receives."

    [[ "$val" == /mnt/* ]]     || _winlab_die "$what is not under /mnt: $val"
    [[ "$val" != *//* ]]       || _winlab_die "$what has an empty path segment: $val"
    [[ "$val" == */ppt-lab ]]  || _winlab_die "$what does not end in /ppt-lab: $val"

    # /mnt/c/Users/<name>/ppt-lab is five segments. Anything shallower means a
    # component came back empty, and a recursive delete there reaches far too
    # much of the Windows drive.
    local depth; depth="$(printf '%s' "$val" | tr -cd '/' | wc -c)"
    [[ "$depth" -ge 4 ]] || _winlab_die "$what is too shallow to clear safely: $val"
}
_winlab_check "$LAB" "the scratch root (LAB)"
unset -f _winlab_check
