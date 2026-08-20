# CodeHighlight — installing it

A PowerPoint add-in that renders source code with VS Code's colours, entirely
offline. You type code into a block on the slide and press **Stylize**.

This file covers getting it onto a machine and set up correctly. **`README.md`
covers using it.**

Decks made with it open and edit correctly on any machine, with or without the
add-in. Someone without it sees exactly what you see — an ordinary rounded
rectangle with ordinary per-character font colours.

What they lose is only the **Stylize button**: nothing will re-colour their
edits automatically. The block is a perfectly ordinary PowerPoint text box, so
they can still select any text in it and change its colour, size or font by
hand exactly as they would anywhere else. Nothing is locked.

---

## 1. What you need

- **PowerPoint 2010 or newer, Windows desktop.** The ribbon markup the add-in
  uses is the binding constraint.
- PowerPoint for the web, and the mobile apps, have no VBA and can never run
  it. They still open and edit the decks it produces.

`tools/check-env.ps1` in the repository reports a machine's version, build,
macro policy and Trusted Locations. It is read-only and safe to send to someone
before installing.

## 2. Install

You need the file itself first. **Download it here:**

<https://github.com/scohen-byte/CodeHighlight/releases/latest/download/CodeHighlight.ppam>

That link always points at the current release, so it stays good. No GitHub
account is needed. (You can also build it yourself — that is **section 5** —
but nobody installing it has any reason to.)

1. Copy **`CodeHighlight.ppam`** into:

   ```
   %APPDATA%\Microsoft\Addins
   ```

   Paste that path into the File Explorer address bar to get there. This folder
   is a Trusted Location on a stock Office install, which is why installing here
   avoids the macro-security prompt entirely.

2. **Right-click the file → Properties → tick Unblock → OK.** Windows marks
   everything that came from the web, and this clears that mark. A Trusted
   Location does bypass the block on its own, but unblocking costs one click
   and removes a whole class of confusion. Skip it only if you built the file
   yourself, in which case there is no mark to clear.

3. In PowerPoint: **File → Options → Add-ins**. At the bottom, set
   **Manage: PowerPoint Add-ins** and click **Go**. Then **Add New…**, pick the
   file, and close.

4. A **Code** tab appears in the ribbon.

If the tab misbehaves after an update, untick and re-tick the add-in in that
same dialog. PowerPoint caches it while loaded.

## 3. Turn off three autocorrect settings

**Do this before teaching with it.** PowerPoint autocorrects as you type, and
three of its habits quietly break code.

**File → Options → Proofing → AutoCorrect Options…**

On the **AutoCorrect** tab:

- Untick **Capitalize first letter of sentences**. Without it, typing `x = 3`
  gives you `X = 3` — a different program.
- In the **Replace text as you type** list, find the entry replacing **`i`**
  with **`I`**, select it and press **Delete**. Without it, `for i in
  range(10)` becomes `for I in range(10)`, and `i` is the most common loop
  variable there is.

  Note this replacement list is shared across Office, so removing it stops Word
  doing the same thing. That is usually welcome, but it is not a
  PowerPoint-only change.

On the **AutoFormat As You Type** tab:

- Untick **"Straight quotes" with "smart quotes"**. Without it, `"hi"` becomes
  `"hi"` with typographic quotes, which is not valid Python.

Stylize repairs the quotes for you, so that one is belt and braces. **It
cannot repair the two capitalisation problems** — nothing can tell a variable
you meant to call `X` from one autocorrect changed, so if either slips through
you have to correct the letter by hand.

## 4. If your IT policy blocks the add-in

Possible on a managed university machine, where macro settings come from Group
Policy and cannot be changed by the user. `check-env.ps1` reports whether the
setting came from policy.

The fix is to sign it. **SelfCert.exe** ships with Office and creates a
self-signed certificate. Sign the VBA project, and the user trusts the publisher
once on first load, which also survives later updates to the file.

If add-ins are blocked outright, your decks still open and edit perfectly on
that machine. Only re-styling is lost. That is a deliberate design goal,
not a workaround.

## 5. Building it from source

Only for changing the add-in. To *use* it you need nothing in this section —
the built file is self-contained, and section 2 is the whole install.

You need **WSL** alongside PowerPoint, since the build drives PowerPoint over
COM from a shell script. From the repository root:

```
tools/build-addin.sh              # builds dist/CodeHighlight.ppam
tools/build-addin.sh --install    # builds, then replaces the installed copy
```

Three things have to be true, and the script says which one is missing rather
than failing obscurely:

- **PowerPoint must be closed.** It holds the add-in open while loaded, and a
  rebuild behind its back is silently discarded. The script names the process
  it found, including a windowless orphan left by a test run — that one is easy
  to miss, because there is nothing on screen to close.
- **VBA project access must be trusted:** File → Options → Trust Center →
  Trust Center Settings → Macro Settings → **Trust access to the VBA project
  object model**. The build imports the modules through that object model, so
  without it there is nothing to import into. `tools/check-env.ps1` reports this
  along with everything else.
- **A Windows profile reachable from WSL.** The build stages files under
  `%USERPROFILE%\ppt-lab`, which it derives per machine — nothing here assumes
  a particular user name.

`--install` keeps the previous add-in beside the new one as
`CodeHighlight.ppam.prev`, so a bad build is one rename away from undone. The
build itself never touches the installed copy.

---

Installed and set up? **`README.md`** has the buttons and the walkthroughs.
