# CodeHighlight

A PowerPoint add-in that renders source code with VS Code's colours, entirely
offline, with no dependency on anything outside PowerPoint itself.

Type code into a block on the slide and press **Stylize**.

Built for teaching: the palette is matched to what a student sees in their
editor, the sizes are chosen to read from the back of a lecture hall, and the
add-in refuses to shrink code past the point of legibility.

`INSTALL.md` covers installation, including two PowerPoint autocorrect settings
that quietly break code and should be turned off first.

## Decks stay portable

The result is an ordinary rounded rectangle with ordinary per-character font
colours. No custom XML parts, no external files, no linked objects. Someone
without the add-in sees exactly what you see and can edit the text normally.

What they lose is only the automatic re-colouring. The block is a plain
PowerPoint text box, so its colours, sizes and fonts can still be changed by
hand like any other text. Nothing is locked.

Sharing a deck therefore needs nothing extra. The font is **Consolas**, which
ships with Windows, so it is present on essentially any machine that can open
the file. On a Mac, or anywhere Consolas is missing, PowerPoint substitutes
another font and the block's alignment shifts — the code stays readable and
editable.

## The buttons

| Button | What it does |
|---|---|
| **New block** | Inserts an empty dark block. The placeholder text is selected, so just start typing |
| **Stylize** | Colours the selected block, refits it, and re-syncs its numbers, guides and notes. The button you press constantly |
| **Stylize all** | The same, for every code block on the slide |
| **Language** | Which lexer to use. Stored per block |
| **Line numbers** | Show or hide the numbers for the selected block. A separate shape, so the code text stays pure source. Off by default |
| **Indent guides** | Show or hide the vertical nesting lines, drawn as real shapes. Off by default |
| **Emphasize** | Band the selected lines and fade the rest |
| **Note** | Attach an explanation to a line of code |
| **Note size** | Font size for notes. Auto sizes each one from its block |
| **Note colour** | One of six presets. The text colour follows the fill |
| **Copy code** | Put the block's code on the clipboard, ready to paste into an editor |
| **Strip** | Back to plain uncoloured text. Stylize brings it all back |
| **Step through** | Build a walkthrough: one slide per line, each emphasising the next |
| **Build up** | The same, with the emphasis growing downward rather than moving |
| **Hide lines** | Cover lines with a `?`, keeping the block's exact layout, so the class can guess them |
| **Reveal next** | Duplicate the slide with nothing hidden — the answer, one click after the question |
| **Size** | The block's size in points. Pick a rung or type any size |
| **Larger / Smaller** | Step the size ladder. Larger stops where the code would stop fitting the slide |
| **Fit** | Pick the largest size that fits. Warns rather than shrinking below readable |

The block grows as you type. Press **Stylize** after editing to re-colour it —
PowerPoint gives an add-in no way to notice that you changed the text.

## Using it

### Emphasising lines

Click into the block, select the lines you want to draw attention to, and press
**Emphasize**. Those lines get a band behind them and everything else fades
back. Press it again with the block itself selected — not text inside it — to
clear.

The choice is stored on the block and survives every Stylize, which is what
makes a walkthrough cheap: duplicate the slide, select the next line, press
Emphasize.

### Building a walkthrough

Select a block and press **Step through**. The slide is duplicated once per line
of code, each copy emphasising the next line, and a final copy is added with
nothing emphasised — so the sequence opens and closes on the whole code and you
simply advance through the slides as you talk.

**Build up** does the same except the emphasis grows downward: each slide shows
everything covered so far rather than one line. Use Step through for tracing
what runs next, Build up for assembling code piece by piece.

Two things worth knowing:

- **Blank lines get no slide of their own.** A slide emphasising an empty line
  is a dead beat.
- **Select text first to walk only part of a block.** Without a selection it
  walks every line, which on a forty-line block means forty slides. Past twelve
  it asks before going ahead.

The generated slides are ordinary slides. Reorder, delete or edit any of them,
and press Stylize on one if you change its code.

### Hiding lines, to ask rather than tell

Select lines inside the block and press **Hide lines**. They disappear behind a
panel carrying a bold white question mark, and the block keeps its exact size —
nothing shifts, so the audience can see precisely how much is missing and where.
Press **Reveal next** and the slide is duplicated with nothing hidden, giving
you the answer slide immediately after the question.

To unhide without adding a slide, select the block itself — not text inside it —
and press **Hide lines** again. **Strip** also clears it.

The code is covered, not deleted. It is still in the file, and anyone who opens
the deck can drag the panel aside. That is fine when you reveal in the room, but
do not use it to withhold answers from a deck you hand out beforehand.

### Notes on a line

Put the cursor on a line and press **Note**. A note appears beside the block
with a thin line pointing at that line of code, and its placeholder text is
selected so you can type straight over it. No need to select anything first —
the cursor is enough.

The note is anchored, not merely placed. Edit the code above it, change the
size, or drag the whole block, and Stylize brings the note back level with its
line. Drag a note somewhere you prefer and Stylize leaves it there.

**On a walkthrough slide the block itself is enough.** With the block selected
rather than text, the note attaches to the emphasised line — so annotating a
generated Step through is one click per slide, with nothing to aim at. Build up
emphasises a range, and the note goes on its **last** line, which is the one
that slide has just reached.

**Note size** and **Note colour** apply to the selected block's notes straight
away, and to every note made afterwards. Auto sizes each note from the block it
belongs to, which is usually what you want. The six colours are fixed rather
than a full picker, because each one is checked against the text colour it gets
— the light preset gets dark words and a faint edge so it does not dissolve into
a white slide. The choice is stored in the deck, not in the add-in, so it
travels with the file.

- **Two notes on nearby lines stack** rather than landing on top of each other.
- **A block too wide to leave a margin** puts its notes below itself instead,
  with the pointer doing the work of saying which line is meant. The slide runs
  out after two or three, and the add-in says so rather than letting one hang
  off the edge. Making the block smaller is usually the better answer.
- **Strip does not remove notes.** Everything else it removes can be rebuilt by
  pressing Stylize, and typed words cannot, so notes are treated as content.
  To clear them, press **Note** with the block itself selected. That one asks
  first.

### Fit will sometimes refuse to be helpful, on purpose

If a snippet only fits at 12pt it tells you so, and says roughly how many slides
it needs at 16pt instead. 16pt is the floor for reading from the back of a
lecture hall. Splitting code across two slides beats shrinking it past
legibility.

## Adding a language

Python is the only language shipped, but nothing outside `src/modLang*.bas`
holds a Python-specific literal. The scanner is generic and driven by a
`LangDef` table: comment markers, quote characters, escape and interpolation
rules, keyword sets, bracket pairs. Adding a language is a new table file and
one `Register` line — the lexer, renderer, palette and ribbon are untouched.

See `PLAN.md` §5b, including an honest note on where that stops working.

## Working on it

```bash
tools/check-vba.sh                  # static checks. ALWAYS run before driving PowerPoint
tools/run-lexer-tests.sh            # the scanner against its reference, whole corpus
tools/run-slice.sh lab_snippet      # render one sample to PNG
tools/run-feature.sh NoteTest band  # drive one modSelfTest function, keep the PNG
tools/build-addin.sh --install      # rebuild and deploy. PowerPoint must be CLOSED
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
