# CodeHighlight

A PowerPoint add-in that renders source code with VS Code's colours, entirely
offline, with no dependency on anything outside PowerPoint itself.

Type code into a block on the slide and press **Stylize**.

Built for teaching: the palette is matched to what a student sees in their
editor, the sizes are chosen to read from the back of a lecture hall, and the
add-in refuses to shrink code past the point of legibility.

## Decks stay portable

The result is an ordinary rounded rectangle with ordinary per-character font
colours. No custom XML parts, no external files, no linked objects. Someone
without the add-in sees exactly what you see and can edit the text normally.

What they lose is only the automatic re-colouring. The block is a plain
PowerPoint text box, so its colours, sizes and fonts can still be changed by
hand like any other text. Nothing is locked.

## What it does

| | |
|---|---|
| **Stylize** | Colours the selected block, refits it, and re-syncs its numbers and guides |
| **Language** | Which lexer to use, stored per block |
| **Line numbers** | A separate shape, so the code text stays pure source |
| **Indent guides** | Vertical lines marking nesting, drawn as real shapes |
| **Larger / Smaller / Fit** | Steps a size ladder. Growth stops where the code stops fitting the slide |
| **Emphasize** | Bands the selected lines and fades the rest |
| **Step through** | Builds a walkthrough: one slide per line, each emphasising the next, opening and closing on the whole code |
| **Build up** | The same, with the emphasis growing downward rather than moving |
| **Copy code** | The source to the clipboard, not a picture of a rectangle |
| **Strip** | Back to plain text. Stylize brings it all back |

`INSTALL.md` covers installation and use, including two PowerPoint autocorrect
settings that quietly break code and should be turned off first.

## Adding a language

Python is the only language shipped, but nothing outside `src/modLang*.bas`
holds a Python-specific literal. The scanner is generic and driven by a
`LangDef` table: comment markers, quote characters, escape and interpolation
rules, keyword sets, bracket pairs. Adding a language is a new table file and
one `Register` line — the lexer, renderer, palette and ribbon are untouched.

See `PLAN.md` §5b, including an honest note on where that stops working.

## Working on it

```bash
tools/check-vba.sh              # static checks. ALWAYS run before driving PowerPoint
tools/run-lexer-tests.sh        # the scanner against its reference, whole corpus
tools/run-slice.sh lab_snippet  # render one sample to PNG
tools/build-addin.sh --install  # rebuild and deploy. PowerPoint must be CLOSED
```

The scanner is tested by diffing it against a reference classifier
(`tools/lexref.py`) that emits one mask character per source character, so a
mistake shows up as a wrong column rather than as a vague sense that the
colours look off. All samples must match.

`check-vba.sh` first is not a nicety: a VBA compile error does not come back as
an error, it opens the VBA editor with a modal dialog and every later COM call
blocks — which from the outside is indistinguishable from a hang.

## Requirements

PowerPoint 2010 or newer, Windows desktop. Development runs from WSL, driving
PowerPoint over COM through `powershell.exe`.
