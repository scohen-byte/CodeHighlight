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

## Two rules that apply everywhere

**Press Stylize after editing.** PowerPoint gives an add-in no way to notice
that you changed the text, so Stylize is the step that re-syncs everything —
colours, numbers, guides, prompts, notes and arrows. It is safe to press as
often as you like.

**You usually need not select the block first.** When a slide holds exactly one
code block, the buttons use it. With two there is no way to guess which you
meant, so they ask. Three buttons are the exception — **Emphasize**, **Hide
lines** and **Arrow** — because for those, "the block is selected rather than
text inside it" *means* clear this, and pressing one with nothing selected would
silently throw the marking away.

---

# The ribbon

The **Code** tab, group by group.

## Block

| | |
|---|---|
| **New block** | Inserts an empty dark block, placeholder selected so you can start typing. The arrow beside it offers **Code block** or **Transcript** |
| **Stylize** | Colours the block, refits it, and re-syncs everything hanging off it. The button you press constantly |
| **Strip** | Back to plain uncoloured text. Stylize brings it all back |
| **Copy code** | The block's source to the clipboard — no prompts, no output |
| **Indent guides** | Vertical lines marking each level of nesting, drawn as real shapes. Off by default |
| **Line numbers** | Numbers in their own shape, so the code text stays pure source. Off by default |
| **First line** | The number the first numbered line gets. Per block |

**Copy code is load-bearing, not a convenience.** In a transcript the block's
text contains prompts, which are not code. Copy code strips them and drops the
output lines, so what you paste is what you would run. That is where the
pure-source guarantee lives.

### Code split across several slides

**First line** sets the number the first numbered line gets. Put 1 on the first
slide's block and 24 on the second, and the numbering carries on instead of
restarting. It is stored per block, survives Stylize, and turns the line numbers
on when you set it — a start number with no numbers showing does nothing.

The gutter widens for the highest number that will actually appear, not for the
line count, so a twelve-line block starting at 98 gets three digits' worth of
room rather than two.

### What the numbers count

In plain code they are **positions in the file**: leading and trailing blank
lines get no number, every interior line does, including blanks. That is so
"look at line 10" finds line 10 in the student's editor.

In a transcript there is no file, so they count **statements** instead — only
the lines you typed. Blank lines and output get none. The two rules differ
because the numbers answer different questions.

## Size

| | |
|---|---|
| **Size** | The block's size in points. Pick a rung or type any size |
| **Larger / Smaller** | Step the size ladder. Larger stops where the code would stop fitting the slide |
| **Fit** | Pick the largest size that fits |

**Fit will sometimes refuse to be helpful, on purpose.** If a snippet only fits
at 12pt it tells you so, and says roughly how many slides it needs at 16pt
instead. 16pt is the floor for reading from the back of a lecture hall.
Splitting code across two slides beats shrinking it past legibility.

## Lines

| | |
|---|---|
| **Emphasize** | Band the selected lines and fade the rest |
| **Arrow** | Point at a line from the left margin, leaving the code at full contrast |

These are two different instruments. **Emphasize** says *ignore everything
else* — it bands one line and fades the rest, which is right when the
surroundings are noise. **Arrow** says *look here, and keep reading*. A short
snippet whose whole point is that it is readable should not have most of it
greyed out to draw the eye to line one.

Emphasize takes the lines you select; press it again with the block itself
selected to clear. The choice is stored on the block and survives every Stylize,
which is what makes a walkthrough cheap.

Arrow takes the line the cursor is on, or — with the block selected — the
emphasised line. Press it again on the same line to remove it, or with the block
selected and nothing emphasised to clear them all. Several lines can carry
arrows at once, and they can be different colours.

The arrow is sized from the line and sits in whatever room there is to the left
of the block, so a block near the slide edge gets a shorter one. Move the block
right and the arrow grows. **Strip** removes arrows, unlike notes: an arrow holds
nothing you typed.

## Transcript

| | |
|---|---|
| **Transcript** | This block is a session at an interpreter |
| **Prompt** | These lines were typed: put the prompt back on them |
| **Output** | These lines were printed: take the prompt off |

Type the **whole session into one block** — what you typed and what it printed,
as plain text. Then put the cursor on a line the interpreter *printed* and press
**Output**. That single press turns the block into a transcript, prompts every
line you typed (`>>>`, with `...` on the body of a statement, nothing on a blank
line), and leaves the line you marked bare.

Press once per run of output; the marking adds up. A bare cursor is enough for
one line. **Prompt** does the reverse, which is how you promote a line you have
just typed — a new line arrives without a prompt, so it counts as output.

**The prompt lives in the text.** It was a separate column at first, on the
argument that the block's text should be pure source — and two shapes that must
agree line for line eventually disagree, at which point the prompts land on top
of the code. In the text there is nothing to keep in step: the prompt indents
the code because it *is* characters in front of the code.

**There is no stored list of output lines**, either. A line carrying a prompt is
one you typed; a bare one is output. The text *is* the record, so editing cannot
put it out of step — insert a line anywhere and nothing moves. That matters
because it was a list of line numbers twice over, and both times inserting a
line near the top made every marking below it point one line too high.

Output lines lose their syntax colour, which is the cue students already know
from a terminal: the interpreter is showing a value, not code. They sit at
column zero while the code is indented, which is what a terminal does — it
indents what you typed and prints its reply hard against the left.

## Notes

| | |
|---|---|
| **Note** | Attach an explanation to a line of code |
| **Output** | Attach what the code printed, dressed as a terminal |
| **Delete** | Remove the note or output you have singled out, and its connector |

Put the cursor on a line and press **Note**. A note appears beside the block with
a thin connector pointing at that line, placeholder selected so you can type
straight over it. No selection needed — the cursor is enough.

**On a walkthrough slide the block itself is enough.** With the block selected
rather than text, the note attaches to the emphasised line, so annotating a
generated Step through is one click per slide. Build up emphasises a range, and
the note goes on its **last** line.

**The connector is a real connector**, attached at one end to an invisible marker
on its line and at the other to the note, so PowerPoint reroutes it as you drag.
Move a note across the slide and the line stretches to follow it, live.

The note is anchored, not merely placed. Edit the code above it, change the size,
or drag the whole block, and Stylize brings it back level with its line. Drag a
note somewhere you prefer and Stylize leaves it there. Two notes on nearby lines
stack rather than landing on top of each other.

**Output** is the same machinery dressed as a terminal — the code font, a darker
panel with an edge, and a green arrow opening the line. Use it when the code is a
program and the result is an aside about one line of it; use **Transcript** when
the slide *is* a session at the interpreter. Long output wraps flush with the
first character after the arrow, so a second line reads as a continuation rather
than as a second result.

**Strip does not remove notes.** Everything else it removes can be rebuilt by
pressing Stylize, and typed words cannot. To clear them all, press **Note** with
the block itself selected — that one asks first.

## Walkthrough

| | |
|---|---|
| **Step through** | One slide per line, each emphasising the next |
| **Build up** | The same, with the emphasis growing downward rather than moving |
| **Point at** | One slide per line, an arrow moving down the margin and no fade |
| **Note per step** | Every generated slide arrives with a note on the line it is about |

Select a block and press one. The slide is duplicated once per line of code, and
for Step through and Point at a final copy is added with nothing marked — so the
sequence opens and closes on the whole code and you simply advance as you talk.

Use **Step through** for tracing what runs next, **Build up** for assembling code
piece by piece, and **Point at** when the snippet is short enough that every line
should stay legible throughout.

The newest line is **bold** in all three. The band changes the background and the
arrow sits outside the block; bold changes the letters, and a slide that has gone
to the trouble of picking out one line wants all of it. Only a walkthrough does
this — an arrow you place by hand leaves its line alone.

Two things worth knowing:

- **Blank lines get no slide of their own.** A slide about an empty line is a
  dead beat.
- **Select text first to walk only part of a block.** Without a selection it
  walks every line, which on a forty-line block means forty slides. Past twelve
  it asks before going ahead.

The generated slides are ordinary slides. Reorder, delete or edit any of them,
and press Stylize on one if you change its code.

## Hide and reveal

| | |
|---|---|
| **Hide lines** | Cover the selected lines with a `?`, keeping the block's exact layout |
| **Reveal next** | Duplicate the slide with one covered region uncovered |

Select lines and press **Hide lines**. They disappear behind a panel carrying a
bold white question mark, and the block keeps its exact size — nothing shifts, so
the audience can see precisely how much is missing and where.

The marking **adds up**, so you can cover several separate regions — every run of
output in a transcript, say — one press per region. **Reveal next** then uncovers
**one** region at a time: the one your cursor is in, or the topmost remaining.
Four covered regions is four questions, and answering them all at the first press
would defeat the point.

To unhide without adding a slide, press **Hide lines** again on a covered region,
or select the block itself to clear them all. **Strip** also clears it.

The code is covered, not deleted. It is still in the file, and anyone who opens
the deck can drag the panel aside. That is fine when you reveal in the room, but
do not use it to withhold answers from a deck you hand out beforehand.

## Appearance

The deck's own settings — the things you touch once and then leave alone. Laid
out like the Home tab's Font group: one word naming the thing, then unlabelled
dropdowns whose contents say what they are.

| | |
|---|---|
| **Notes** | Colour, font and size for notes |
| **Arrow** | Colour for margin arrows |
| **Language** | Which lexer to use. Stored per block |

**Notes.** The three apply to the note you have singled out — the one on your
cursor's line, or the one you have clicked into — and each changes only what it
is named after, so two notes on one slide can differ in all three. With nothing
singled out they only record the choice for notes made afterwards.

Twelve colours, each shown as a swatch of itself. A fixed list rather than a
picker, because every one is checked against the text colour it gets: the test
asserts all twelve clear 4.5:1, the worst being Emerald at 4.57. The first six
are quiet, for a note that is an aside about the code; the rest are saturated,
for one that should be the loudest thing on the slide. The text colour is
whichever of light or dark reads better on the fill, so amber gets dark words and
crimson gets white ones.

Size is 24pt to begin with; **Auto** sizes each note from its block instead.
**Default font** inherits your presentation's own body font, which is the default
and usually right — it cannot be applied *back* to a note that already has a
named font, since no font name means "inherit".

**Arrow.** Eight colours: the syntax hues, darkened. Taking the palette as it
stands does not work — those colours are chosen against the block's `#1F1F1F`,
and an arrow sits on the slide. Measured against white, the teal of a class name
manages 2.04:1 and the purple of a keyword 2.78:1. Each keeps its hue and loses
enough lightness to clear 4.5:1 on white; the weakest is Gold at 4.92.

It applies to the arrow you have singled out, so two arrows on a slide can
differ — *this is the bug* beside *this is the fix*. With nothing singled out it
only records the colour, so the next arrow gets it, and so does every arrow a
walkthrough makes — which is why **Point at** comes out uniform without being
told to.

---

## Adding a language

Python is the only language shipped, but nothing outside `src/modLang*.bas`
holds a Python-specific literal. The scanner is generic and driven by a
`LangDef` table: comment markers, quote characters, escape and interpolation
rules, keyword sets, bracket pairs, and the interpreter's prompts. Adding a
language is a new table file and one `Register` line — the lexer, renderer,
palette and ribbon are untouched.

See `PLAN.md` §5b, including an honest note on where that stops working.

## Working on it

```bash
tools/check-vba.sh                  # static checks. ALWAYS run before driving PowerPoint
tools/run-lexer-tests.sh            # the scanner against its reference, whole corpus
tools/run-slice.sh lab_snippet      # render one sample to PNG
tools/run-feature.sh NoteTest band  # drive one modSelfTest function, keep the PNG
tools/run-feature.sh EverythingTest lab_snippet   # every feature on one block
tools/build-addin.sh --install      # rebuild and deploy. PowerPoint must be CLOSED
```

The scanner is tested by diffing it against a reference classifier
(`tools/lexref.py`) that emits one mask character per source character, so a
mistake shows up as a wrong column rather than as a vague sense that the
colours look off. All samples must match.

`check-vba.sh` first is not a nicety: a VBA compile error does not come back as
an error, it opens the VBA editor with a modal dialog and every later COM call
blocks — which from the outside is indistinguishable from a hang. Three of its
checks exist because of specific failures that reached a build:

- **`check-refs.py`** resolves every `modX.Member` against what `modX.bas`
  exports. VBA compiles one procedure at a time, so a reference to something
  that has moved survives a full green test run and fails the first time a user
  reaches that one procedure.
- **`check-modules.py`** checks that each harness imports every module its
  imports reach. A missing one is reported by VBA as "Variable not defined"
  against a module name, which reads as a bug in the code being tested and is
  not.
- **the dangling-`GoTo` check**, because "Label not defined" takes down a whole
  module and is invisible until that one procedure is called.

`EverythingTest` is the one that catches interaction bugs: eight kinds of child
shape — gutter, guides, band, cover, note, leader, anchor, arrow — competing for
one group and one z-order, styled twice. Equal part counts across the second
pass is the property that matters, since Stylize is the button you press
constantly and anything created rather than repositioned piles up.

## Requirements

PowerPoint 2010 or newer, Windows desktop. Development runs from WSL, driving
PowerPoint over COM through `powershell.exe`.
