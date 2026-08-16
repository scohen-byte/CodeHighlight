# CodeHighlight — installing and using it

A PowerPoint add-in that renders source code with VS Code's colours, entirely
offline. You type code into a block on the slide and press **Highlight**.

Decks made with it open and edit correctly on any machine, with or without the
add-in. Someone without it sees exactly what you see — an ordinary rounded
rectangle with ordinary per-character font colours.

What they lose is only the **Highlight button**: nothing will re-colour their
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

1. Copy **`CodeHighlight.ppam`** into:

   ```
   %APPDATA%\Microsoft\Addins
   ```

   Paste that path into the File Explorer address bar to get there. This folder
   is a Trusted Location on a stock Office install, which is why installing here
   avoids the macro-security prompt entirely.

2. If the file arrived by email, chat or download: **right-click it →
   Properties → tick Unblock → OK.** A Trusted Location does bypass the
   mark-of-the-web block, but unblocking costs one click and removes a whole
   class of confusion.

3. In PowerPoint: **File → Options → Add-ins**. At the bottom, set
   **Manage: PowerPoint Add-ins** and click **Go**. Then **Add New…**, pick the
   file, and close.

4. A **Code** tab appears in the ribbon.

If the tab misbehaves after an update, untick and re-tick the add-in in that
same dialog. PowerPoint caches it while loaded.

## 3. Turn off two autocorrect settings

**Do this before teaching with it.** PowerPoint autocorrects as you type, and
two of its habits quietly break code.

**File → Options → Proofing → AutoCorrect Options…**

- On the **AutoCorrect** tab, untick **Capitalize first letter of sentences**.
  Without this, typing `x = 3` gives you `X = 3` — a different program, and one
  nothing can repair afterwards, because there is no way to tell a variable you
  meant to call `X` from one autocorrect changed.
- On the **AutoFormat As You Type** tab, untick
  **"Straight quotes" with "smart quotes"**. Without this, `"hi"` becomes `"hi"`
  with typographic quotes, which is not valid Python.

Highlight repairs the quotes for you, so this one is belt and braces. It cannot
repair the capitalisation.

## 4. Using it

| Button | What it does |
|---|---|
| **New block** | Inserts an empty dark block. The placeholder text is selected, so just start typing |
| **Highlight** | Colours the selected block, refits it, and re-syncs its line numbers and guides. The button you press constantly |
| **Highlight all** | The same, for every code block on the slide |
| **Language** | Which lexer to use. Stored per block |
| **Line numbers** | Show or hide the numbers for the selected block. Off by default |
| **Indent guides** | Show or hide the vertical nesting lines. Off by default |
| **Emphasise** | Band the selected lines and fade the rest. See below |
| **Copy code** | Put the block's code on the clipboard, ready to paste into an editor |
| **Strip** | Back to plain uncoloured text. Highlight brings it all back |
| **Size** | The block's size in points. Pick a rung or type any size |
| **Larger / Smaller** | Step the font size. Larger stops where the code would stop fitting the slide |
| **Fit** | Pick the largest size that fits. Warns rather than shrinking below readable |

The block grows as you type. Press **Highlight** after editing to re-colour it —
PowerPoint gives an add-in no way to notice that you changed the text.

**Emphasising lines, for stepping through code.** Click into the block, select
the lines you want to draw attention to, and press **Emphasise**. Those lines
get a band behind them and everything else fades back. Press it again with the
block itself selected — not text inside it — to clear.

The choice is stored on the block and survives every Highlight, which is what
makes a walkthrough cheap: duplicate the slide, select the next line, press
Emphasise.

**Fit will sometimes refuse to be helpful, on purpose.** If a snippet only fits
at 12pt it tells you so, and says roughly how many slides it needs at 16pt
instead. 16pt is the floor for reading from the back of a lecture hall. Splitting
code across two slides beats shrinking it past legibility.

## 5. If your IT policy blocks the add-in

Possible on a managed university machine, where macro settings come from Group
Policy and cannot be changed by the user. `check-env.ps1` reports whether the
setting came from policy.

The fix is to sign it. **SelfCert.exe** ships with Office and creates a
self-signed certificate. Sign the VBA project, and the user trusts the publisher
once on first load, which also survives later updates to the file.

If add-ins are blocked outright, your decks still open and edit perfectly on
that machine. Only re-highlighting is lost. That is a deliberate design goal,
not a workaround.

## 6. Sharing a deck

Nothing extra to do. The code block is a normal shape with normal text, so the
deck is self-contained.

The font is **Consolas**, which ships with Windows, so it is present on
essentially any machine that can open the deck. On a Mac, or anywhere Consolas
is missing, PowerPoint substitutes another font and the block's alignment will
shift — the code remains readable and editable.
