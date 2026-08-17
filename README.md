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
| **Language** | Which lexer to use. Stored per block |
| **Line numbers** | Show or hide the numbers for the selected block. A separate shape, so the code text stays pure source. Off by default |
| **Indent guides** | Show or hide the vertical nesting lines, drawn as real shapes. Off by default |
| **Transcript** | Style the block as a session at an interpreter — prompts on every line you typed |
| **Output** | Take the prompt off the selected lines: these were printed, not typed |
| **Prompt** | Put it back: these were typed. A newly typed line starts without one |
| **First line** | The number the first line gets, for code split across slides. Per block |
| **Emphasize** | Band the selected lines and fade the rest |
| **Arrow** | Point at a line from the left margin, leaving the code at full contrast |
| **Arrow color** | Eight syntax hues, darkened to read on a white slide. Deck-wide |
| **Note** | Attach an explanation to a line of code |

| **Output** | Attach what the code printed beside the block, as a note |
| **Delete** | Remove the note or output you have singled out, and its connector |
| **Note size** | Font size for notes. 24pt, or Auto to size each one from its block |
| **Note color** | Twelve presets, each shown as its own swatch |
| **Note font** | Deck default, or one of four that ship with Windows |
| **Copy code** | Put the block's code on the clipboard, ready to paste into an editor |
| **Strip** | Back to plain uncoloured text. Stylize brings it all back |
| **Step through** | Build a walkthrough: one slide per line, each emphasising the next |
| **Build up** | The same, with the emphasis growing downward rather than moving |
| **Note per step** | Every generated slide arrives with a note on the line it emphasises |
| **Bold the line** | Render the newest emphasised line in bold as well as banded |
| **Arrow, not fade** | Mark each generated slide with an arrow instead of fading the rest |
| **Hide lines** | Cover lines with a `?`, keeping the block's exact layout, so the class can guess them |
| **Reveal next** | Duplicate the slide with nothing hidden — the answer, one click after the question |
| **Size** | The block's size in points. Pick a rung or type any size |
| **Larger / Smaller** | Step the size ladder. Larger stops where the code would stop fitting the slide |
| **Fit** | Pick the largest size that fits. Warns rather than shrinking below readable |

The block grows as you type. Press **Stylize** after editing to re-colour it —
PowerPoint gives an add-in no way to notice that you changed the text.

**You usually need not select the block first.** When a slide holds exactly one
code block, the buttons use it. With two there is no way to guess which you
meant, so they ask.

Three buttons are the exception — **Emphasize**, **Hide lines** and **Arrow** —
because for those, "the block is selected rather than text inside it" *means*
clear this. Pressing one with nothing selected would silently throw the marking
away, so they ask you to select first.

## Using it

### Emphasising lines

Click into the block, select the lines you want to draw attention to, and press
**Emphasize**. Those lines get a band behind them and everything else fades
back. Press it again with the block itself selected — not text inside it — to
clear.

The choice is stored on the block and survives every Stylize, which is what
makes a walkthrough cheap: duplicate the slide, select the next line, press
Emphasize.

### Code split across several slides

**First line** sets the number the first numbered line gets. Put 1 on the first
slide's block and 24 on the second, and the numbering carries on instead of
restarting. It is stored per block, survives Stylize, and turns the line numbers
on when you set it — a start number with no numbers showing does nothing.

The gutter widens for the highest number that will actually appear, not for the
line count, so a twelve-line block starting at 98 gets three digits' worth of
room rather than two.

### Pointing at a line without fading the rest

**Emphasize** and **Arrow** are two different instruments. Emphasize says
*ignore everything else* — it bands one line and fades the rest, which is right
when the surroundings are noise. **Arrow** puts a block arrow in the left margin
and leaves the code at full contrast: *look here, and keep reading*. A short
snippet whose whole point is that it is readable should not have most of it
greyed out to draw the eye to line one.

It takes the line the cursor is on, or — with the block itself selected — the
emphasised line. Press it again on the same line to take the arrow away, or with
the block selected and nothing emphasised to clear them all. Several lines can
carry arrows at once.

The arrow is sized from the line and sits in whatever room there is to the left
of the block, so a block near the slide edge gets a shorter one. Move the block
right and the arrow grows. **Strip** removes arrows, unlike notes: an arrow holds
nothing you typed, and one click puts it back.

**Arrow color** offers the syntax hues, darkened. Taking the palette as it
stands does not work — those colours are chosen against the block's `#1F1F1F`,
and an arrow sits on the slide. Measured against white, the teal of a class name
manages 2.04:1 and the purple of a keyword 2.78:1; as text on the block they are
right, as a solid shape on a white slide they look faded, which is the opposite
of what an arrow is for. Each keeps its hue and loses enough lightness to clear
4.5:1 on white, which the test asserts — the weakest is Gold at 4.92.

**Arrow color** applies to the arrow you have singled out — the one your cursor's
line carries, or the one you have clicked into — exactly like Note colour. Two
arrows on a slide can differ, which is the point: *this is the bug* beside *this
is the fix*.

With nothing singled out it only records the colour, so the next arrow gets it —
and so does every arrow a walkthrough makes, which is why **Point at** comes out
uniform without being told to. That is the distinction: a walkthrough's arrows
are one marker moving down a deck and should match, while an arrow you place by
hand is an annotation and need not.

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
tools/run-feature.sh EverythingTest lab_snippet   # every feature on one block
tools/build-addin.sh --install      # rebuild and deploy. PowerPoint must be CLOSED
```

The scanner is tested by diffing it against a reference classifier
(`tools/lexref.py`) that emits one mask character per source character, so a
mistake shows up as a wrong column rather than as a vague sense that the
colours look off. All samples must match.

`check-vba.sh` first is not a nicety: a VBA compile error does not come back as
an error, it opens the VBA editor with a modal dialog and every later COM call
blocks — which from the outside is indistinguishable from a hang. It also
resolves every `modX.Member` reference against what `modX.bas` exports, because
VBA compiles one procedure at a time and a reference to something that has moved
survives a full green test run, then fails the first time a user reaches that
one procedure.

`EverythingTest` is the one that catches interaction bugs: eight kinds of child
shape — gutter, guides, band, cover, note, leader, anchor, arrow — competing for
one group and one z-order, styled twice. Equal part counts across the second
pass is the property that matters, since Stylize is the button you press
constantly and anything created rather than repositioned piles up.

## Requirements

PowerPoint 2010 or newer, Windows desktop. Development runs from WSL, driving
PowerPoint over COM through `powershell.exe`.
