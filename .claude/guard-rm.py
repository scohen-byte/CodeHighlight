#!/usr/bin/env python3
"""Refuse recursive deletes that cannot be read at face value.

A PreToolUse hook sees the command TEXT, before the shell expands anything.
It cannot know what "$STAGE" will become - which is the whole point: on
2026-09-03 a mis-derived "$STAGE" sent `rm -rf` across the entire C: drive.
Nothing was lost, only because Windows denied every file and a timeout killed
it inside Program Files. The path was never verified because it could not be.

So this refuses the SHAPE of the command rather than the target:

  1. A recursive rm whose target is a variable or a command substitution.
     Deleting something you cannot see the name of is the failure mode.
  2. A recursive rm anywhere under /mnt - the Windows filesystem, where the
     irreplaceable things live. The repository is on the WSL side; scratch
     space belongs in mktemp or the session scratchpad.
  3. A recursive rm of /, ~, or $HOME.

Non-recursive rm is untouched: deleting one named file is not this hazard.
Deletes INSIDE project scripts are invisible here and stay that way - the
hook only sees commands issued directly, which is where the mistake was made.
"""
import json
import shlex
import sys

DENY_EXACT = {"/", "~", "$HOME", "${HOME}", "/mnt", "/mnt/"}


def segments(command):
    """Split a command line into pipeline/list segments, crudely but safely."""
    out, buf, i = [], [], 0
    tokens = command.replace("\n", ";").split(";")
    for t in tokens:
        for part in t.split("&&"):
            for piece in part.split("||"):
                out.extend(piece.split("|"))
    return [s.strip() for s in out if s.strip()]


def is_recursive_rm(tokens):
    if not tokens:
        return False
    head = tokens[0].strip("'\"")
    if head != "rm" and not head.endswith("/rm"):
        return False
    for t in tokens[1:]:
        if t == "--":
            break
        if t == "--recursive":
            return True
        if t.startswith("-") and not t.startswith("--"):
            if "r" in t[1:] or "R" in t[1:]:
                return True
    return False


def targets_of(tokens):
    out, after_ddash = [], False
    for t in tokens[1:]:
        if t == "--":
            after_ddash = True
            continue
        if not after_ddash and t.startswith("-"):
            continue
        out.append(t)
    return out


def verdict(command):
    for seg in segments(command):
        try:
            tokens = shlex.split(seg, posix=False)
        except ValueError:
            continue
        if not is_recursive_rm(tokens):
            continue
        for raw in targets_of(tokens):
            t = raw.strip("'\"")
            if "$" in t or "`" in t:
                return ("a recursive rm whose target is unexpanded: %s\n"
                        "This hook sees the command before the shell expands it, so "
                        "what that becomes cannot be checked - and an rm -rf on a "
                        "mis-derived path is what this guard exists to stop.\n"
                        "Write the literal path, or let a project script do the "
                        "delete." % raw)
            if t in DENY_EXACT or t.rstrip("/") in DENY_EXACT:
                return "a recursive rm of %s, which is never intended." % raw
            if t == "/mnt" or t.startswith("/mnt/"):
                return ("a recursive rm under /mnt (%s).\n"
                        "That is the Windows filesystem - the irreplaceable side. "
                        "Use mktemp -d or the session scratchpad for scratch space, "
                        "or run the project script that owns that directory."
                        % raw)
    return None


def main():
    try:
        payload = json.load(sys.stdin)
    except Exception:
        sys.exit(0)
    command = (payload.get("tool_input") or {}).get("command") or ""
    reason = verdict(command)
    if reason:
        json.dump({"hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": "deny",
            "permissionDecisionReason": "Blocked by .claude/guard-rm.py: " + reason,
        }}, sys.stdout)
    sys.exit(0)


if __name__ == "__main__":
    main()
