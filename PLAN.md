# CodeHighlight — implementation plan

A PowerPoint add-in that renders source code with VS Code Dark+ syntax colors,
entirely offline, with no dependency on anything outside PowerPoint itself.

Python is the only language shipped in v1 and the only one being built for now.
The architecture is language-neutral throughout, so a second language is a data
table rather than a rewrite. See section 5b.

Status: Phases −1 through 4 complete. The add-in is built, installed and in use.
Phase 5 (signing and a clean install on a second machine) is what remains, and
it needs Sara. Written 2026-08-14, renamed from PyCodeHighlight on 2026-08-15.

---

## 1. Decisions already made

| Question | Decision |
|---|---|
| Language support | Python in v1. Generic scanner plus a per-language table, so more can be added without touching the lexer, renderer or palette |
| Technology | VBA, packaged as a `.ppam` add-in |
| Platform for the add-in | Windows PowerPoint 2016+ |
| Platform for the *output* | Anywhere. Must open, display and edit correctly with no add-in installed |
| Where you type the code | Directly into the shape on the slide, then press **Highlight** |
| Appearance | Dark filled block, VS Code Dark+ palette |
| Line numbers | Optional, toggled per block, rendered as a separate aligned shape |
| Themes | One (Dark+). Palette kept in its own module so more can be added later |
| Font | Chosen by bake-off in Phase −1, then **embedded in the deck** |
| First step | A visual mockup lab. Nothing else starts until the look is signed off |

### The core idea

**The shape's text is the source.** Nothing is stored anywhere else.

Highlighting reads `shape.TextFrame.TextRange.Text`, tokenizes it, and applies
colors to character ranges. Editing is just normal PowerPoint text editing.
Re-highlighting is idempotent — it resets everything to the default color and
re-applies from scratch, so it never accumulates corruption.

This is what makes the portability requirement free. The result is an ordinary
rectangle with an ordinary text box inside it and per-character font colors.
A colleague without the add-in sees exactly what you see, and can edit the text
normally. They only lose the ability to re-color after editing.

---

## 2. Portability checklist

Every one of these is a hard constraint on the implementation.

- **Font: embedded, not assumed.** See section 2a. The add-in turns on font
  embedding for the presentation, so a good font can be used without requiring
  the reader to have it installed.
- **No custom XML parts, no external files, no linked objects.** Everything the
  reader needs is in the shape itself.
- **Line numbers live in a second plain text box**, not inside the code text.
  The code shape's text stays pure Python at all times.
- **Word wrap off** on the code shape. Autofit off. Both would misalign the
  gutter, and wrapped code is unreadable anyway.
- Shape tags are used for add-in bookkeeping only. If they are missing or
  stripped, the add-in falls back to treating any selected text box as code.
  Nothing about *display* depends on them.

---

## 2a. Font embedding

The visual matters more than font portability, and embedding resolves the
tension. PowerPoint can embed the font into the `.pptx`, so the block renders
correctly on a machine that has never seen it.

**Rules that constrain the font choice:**

- **TrueType only.** `.otf` fonts with PostScript/CFF outlines cannot be
  embedded by PowerPoint. Every candidate below is `.ttf`.
- **Licence must permit embedding** (`fsType` bits in the OS/2 table). SIL OFL
  fonts qualify. Verify with `otfinfo -i <font>.ttf | grep -i embed` or by
  checking whether PowerPoint greys out the option.
- **Embed all characters, not "only the characters used".** The size-saving
  option leaves the deck readable but not properly *editable* on a machine
  lacking the font, which violates the core requirement. Cost of embedding a
  full weight is roughly 250–400 KB. Irrelevant.
- **No-ligature variant.** Per-character coloring splits ligatures anyway, and
  PowerPoint renders them inconsistently. Use JetBrains Mono **NL**, Iosevka
  **Fixed**, or Cascadia **Mono** (not Cascadia *Code*).
- **Mac** PowerPoint reads embedded fonts in recent Microsoft 365 builds. Older
  versions ignore them and substitute. Acceptable, and untestable here.

**The add-in should set embedding automatically.** It is a per-presentation
setting that is trivial to forget, and forgetting it silently breaks the
handoff. Set it when a code block is inserted.

Mechanism to verify in Phase 0: `Presentation.SaveAs` accepts an
`EmbedTrueTypeFonts` argument. If that turns out not to be settable
programmatically in a useful way, fall back to a ribbon button that prompts, and
document the manual path (File > Options > Save > Embed fonts in the file).

### Measured font metrics

Read directly from the `.ttf` files with fontTools on 2026-08-14.

| Font | advance/em | x-height/em | **x-ht ÷ adv** | fsType | Where |
|---|---|---|---|---|---|
| Consolas | 0.550 | 0.490 | 0.89 | Editable | Windows system |
| Cascadia Mono | 0.586 | 0.518 | 0.88 | Installable | Windows system |
| JetBrains Mono NL | 0.600 | 0.550 | 0.92 | Installable | **Per-user only** |
| Iosevka Fixed | 0.500 | ~0.520 | ~1.04 | Installable | not installed |

**Advance width alone is the wrong metric.** Legibility from the back of a room
tracks *x-height*, not point size, so the honest comparison is characters per
line at equal perceived size — the `x-height ÷ advance` column. The three
installed fonts land within 5% of each other. Consolas is narrower, but its
x-height is smaller in the same proportion, so it must be set larger to read
equally well and the density advantage cancels out.

Confirmed empirically in the lab:

| Variant | chars/line | x-height | lines |
|---|---|---|---|
| Consolas 20pt | 73 | 9.8pt | 16 |
| Cascadia Mono 20pt | 69 | 10.4pt | 16 |
| JetBrains Mono NL 20pt | 67 | 11.0pt | 16 |
| **Consolas 22pt** | **67** | **10.8pt** | 15 |

Consolas at 22pt is indistinguishable from JetBrains Mono NL at 20pt on both
axes. So the density argument does not constrain the choice, and the decision
falls to familiarity.

**Leading candidate: Consolas at 20–22pt.** It is VS Code's default
`editor.fontFamily` on Windows, which is literally what a student in the course
sees on screen. It is on every Windows machine, so embedding becomes a
belt-and-braces measure rather than a dependency. Iosevka is the only genuine
density win at ~1.04, but it resembles nothing a student has used, which defeats
the purpose in a teaching context.

---

## 2b. Phase −1 — the visual lab (do this first)

**Nothing else starts until the block looks good.** The add-in is worthless if
the output is ugly, and the look is cheap to iterate on and expensive to retrofit.

The lab is a Python script, run entirely from WSL, that emits a real `.pptx`
using `python-pptx` and `pygments`. Not an HTML approximation — an actual deck
with real per-run font colors that opens in PowerPoint and can be judged
honestly.

```
tools/lab.py  ->  dist/lab.pptx  ->  open in PowerPoint, pick a slide
```

One slide per variant, same Python snippet on every slide, labelled with its
parameters. Axes to vary:

| Axis | Values to try |
|---|---|
| Font | JetBrains Mono NL, Cascadia Mono, Consolas, Iosevka Fixed |
| Size | 12, 14, 16 pt |
| Line spacing | 0.9, 1.0, 1.15 |
| Padding | tight vs generous internal margins |
| Corner radius | square, small, rounded |
| Palette | VS Code Dark+, and one slightly desaturated variant for projectors |

Use a realistic snippet, not `hello world`. Something 20–30 lines with a class,
a decorator, a docstring, an f-string and a comment, so the palette is judged on
the code you actually present.

**Two things make this loop fast:**

1. `pygments` does the tokenizing, so no lexer is needed to evaluate the look.
   The VBA lexer only has to *reproduce* the winning classification later.
2. PowerPoint COM is reachable from WSL, so slides can be exported to PNG
   (`$slide.Export(...)`) and inspected directly without a manual round trip.

**Output of this phase is a spec**, not just a preference: exact font name, point
size, line spacing, margins, fill color, corner radius, and the final hex palette.
Section 6 gets rewritten with those numbers, and the VBA reproduces them.

The lab script is kept. It stays useful as a reference renderer for checking the
VBA output against a known-good result.

### Layout rules discovered by building the lab

These are bugs found by looking at rendered output, and the VBA renderer must
reproduce every one of them.

- **Autoshapes default to centred text.** Every code paragraph must have its
  alignment set to left explicitly, or the block renders ragged and centred.
  This is not visible until you look at a rendered slide.
- **The gutter goes inside the dark block**, not beside it. Implemented by
  giving the code text frame a left margin of `padding + gutter width` and
  overlaying a right-aligned text box on the block's left edge. A gutter sitting
  outside the block on the white slide looks nothing like an editor.
- **The block hugs its content vertically.** Height is
  `line_count × line_spacing + 2 × padding`. A fixed-height block leaves dead
  dark space under short snippets.
- **Line spacing is set in exact points, not as a multiple.** This makes the
  gutter align deterministically instead of depending on the font's internal
  line metrics, which is the single biggest simplification available to the VBA
  gutter code.
- At 24pt the block holds roughly **13 lines × 61 characters**. That is the real
  budget for a classroom slide, and it means long examples get split across
  slides rather than shrunk.

---

## 2c. Tooling and the WSL / Windows split

Verified working on 2026-08-14:

| Check | Result |
|---|---|
| `powershell.exe` callable from WSL | yes |
| PowerPoint COM automation via interop | yes, version 16.0 |
| `python-pptx` | 1.0.2 |
| `pygments` | 2.20.0 |

**Stay in WSL. Do not move to native Windows.** Interop covers everything the
add-in work needs:

- Drive PowerPoint with `powershell.exe -NoProfile -Command '...'` using
  `New-Object -ComObject PowerPoint.Application`.
- Run a macro in the loaded add-in with `$ppt.Run("MacroName")`, which makes the
  VBA testable from a script instead of by hand.
- Export a slide to PNG with `$slide.Export(path, "PNG", w, h)` for visual
  inspection.
- Translate paths with `wslpath -w`.

**One rule:** the build script stages output into a `/mnt/c/...` directory.
Driving COM against `\\wsl$\...` paths is slower and occasionally unreliable.
Keep the repo in WSL, put the artifacts on the C: drive.

---

## 3. Repository layout

Source of truth is this repo, under WSL. The `.bas` files are imported into the
VBA editor on the Windows side.

```
addin/
  PLAN.md                     this file
  INSTALL.md                  written in Phase 5, ships with the .ppam
  src/
    modRibbon.bas             ribbon callbacks
    modLangRegistry.bas       LangDef type + the list of languages
    modLangPython.bas         the Python table. Data only, no logic
    modLexer.bas              generic scanner, driven by a LangDef  <- the hard part
    modTheme.bas              color table, keyed by language-neutral token class
    modRender.bas             applies spans to a shape
    modBlock.bas              create / find / identify code blocks
    modGutter.bas             line number shape
    modSelfTest.bas           inserts the test samples onto a scratch slide
  ribbon/
    customUI14.xml            ribbon tab definition
  tools/
    lab.py                    Phase -1 visual lab, python-pptx + pygments
    pack-ribbon.sh            injects the ribbon XML into the .ppam
    ppt.ps1                   COM helpers, called via powershell.exe interop
    check-env.ps1             read-only readiness check, safe to send to others
  tests/samples/
    *.py                      lexer edge cases, and the lab snippet
  dist/
    lab.pptx                  generated, staged to /mnt/c for PowerPoint
    CodeHighlight.ppam        the shippable artifact
```

### Development round trip

1. Edit `.bas` files here in WSL.
2. In PowerPoint's VBA editor: remove the old module, then File > Import File.
   Importing without removing first creates `modLexer1`, which is a confusing
   way to lose an afternoon.
3. Test on a slide.
4. Save the `.ppam`.
5. Run `tools/pack-ribbon.sh` if the ribbon XML changed.

The `.ppam` lives on the Windows side. Reach it from WSL via `/mnt/c/...`.

---

## 4. The ribbon

`.ppam` add-in macros do not appear in PowerPoint's Macros dialog, so a custom
ribbon tab is the only sane way to expose the commands. A `.ppam` is a zip
package. The ribbon is added by injecting `customUI/customUI14.xml` and a
relationship entry into `_rels/.rels`, which `tools/pack-ribbon.sh` does.

Tab **Code**, three groups. Written and verified as XML in
`ribbon/customUI14.xml`.

| Group | Control | Action |
|---|---|---|
| Code block | New block | Insert a pre-formatted empty dark block at the default size |
| Code block | Language | Dropdown, populated from `modLangRegistry`. Sets the language tag on the selected block |
| Highlight | Highlight | Tokenize and color the selected block. This is the button you press constantly |
| Highlight | Highlight all | Same, for every code block on the current slide |
| Highlight | Line numbers | Toggle button. Gutter on or off for the selected block |
| Size | Larger / Smaller | Step the block up or down the size ladder. Larger stops where content stops fitting |
| Size | Fit | Auto-pick the largest size that fits. Warns if that lands below 16pt |

The language dropdown is the one piece of UI that exists purely for
extensibility. It is populated by callbacks that read `modLangRegistry`, so
adding a language never touches the ribbon XML. With one language registered it
shows a single entry, which costs nothing and keeps the wiring honest from day
one rather than being retrofitted.

Icons come from Office's built-in `imageMso` set, so no image assets are needed.
An `imageMso` id that does not exist renders blank — it does not stop the tab
loading — so check every icon during the Phase 0 spike and swap any that come up
empty.

Ribbon callbacks must have the signature `Sub Name(control As IRibbonControl)`,
with extra arguments for dropdowns and toggles. A wrong signature fails at
*click* time with "Cannot run the macro", which looks exactly like a packaging
problem and is not.

---

## 5. The lexer — `modLexer.bas`

This is 80% of the work and 100% of the risk. Everything else is plumbing.

A hand-written character scanner, not regular expressions. `VBScript.RegExp` is
available but a scanner is faster, handles nesting and multi-line strings
correctly, and is far easier to debug.

**The scanner never names a language.** It is parameterised by a `LangDef`
record — comment markers, quote characters, escape rules, keyword sets — and
Python is one such record in `modLangPython.bas`. Section 5b covers why, and
what it costs.

**Input:** the raw string from `TextRange.Text`, plus a `LangDef`.
**Output:** an array of spans — `(start, length, tokenClass)`, 1-based start.

### Text encoding gotchas

- PowerPoint separates paragraphs with **CR (`vbCr`, chr 13)**, not CRLF.
- A soft line break (Shift+Enter) is **vertical tab (chr 11)**.
- Treat both as end-of-line for comment termination.
- `TextRange.Characters(start, length)` is **1-based**. Off-by-one here produces
  colors shifted by one character across the whole block, which looks bizarre
  and is easy to misdiagnose.

### Scanner states

Each state reads its trigger characters from the `LangDef` rather than from a
literal. The Python values are given as examples.

1. **Line comment** — any marker in `LineComments` (`#`) outside a string, runs
   to end of line.
2. **Block comment** — `BlockCommentOpen` … `BlockCommentClose`. Skipped
   entirely when the open marker is empty, as it is for Python.
3. **String** — any char in `QuoteChars` (`'` and `"`), plus triples when
   `TripleQuotes` is set.
   - Optional 1–2 char prefix from `StringPrefixes`, case-insensitive:
     `r b u f rb br fr rf`.
   - `EscapeChar` consumes the next character, *unless* the prefix contains a
     letter from `RawPrefixChars`.
   - Triple-quoted strings span lines and are the most common source of lexer
     bugs. A `#` or a lone `"` inside one must not terminate anything.
4. **Number** — decimal, float, exponent, the `NumberPrefixes` (`0x` `0o` `0b`),
   `DigitSep` (`_`), and the `NumberSuffixes` (`j`).
5. **Identifier** — `[A-Za-z_][A-Za-z0-9_]*`, then classified (below).
6. **Decorator** — `DecoratorChar` (`@`) as the first non-whitespace on a line,
   plus the dotted name that follows. Skipped when the char is empty.
7. **Everything else** — operators and punctuation, default color.

### Identifier classification

Checked in this order. Every set named here is a field on the `LangDef`, so the
order is the general rule and the contents are the per-language part.

1. `ControlKeywords` → purple
2. `DeclKeywords` → blue
3. `SelfWords` (`self` / `cls`) → blue
4. Immediately preceded by a `FuncDefKeywords` word (`def`) → yellow
5. Immediately preceded by a `TypeDefKeywords` word (`class`) → teal
6. In `TypeNames` → teal
7. Immediately followed by `(` → yellow
8. In `Builtins` → yellow
9. Otherwise → default

Word sets are expanded from their space-delimited table entries into
`Collection` objects **once per block**, not per token, by
`modLangRegistry.WordSet`. Membership is then a keyed lookup. Never test with
`InStr` on a keyword blob — it matches substrings, so `information` comes back
as the keyword `in`.

### Explicitly out of scope for v1

- f-string interpolation. `f"hi {name}"` colors as one string. Correct
  highlighting requires recursively lexing the `{...}` regions. Deferred to
  Phase 6.
- Type annotations get no special treatment beyond normal identifier rules.
- Semantic awareness of any kind. This is a lexer, not a parser.

---

## 5b. Extensibility to other languages

Only Python is being built. But the cost of keeping the design language-neutral
is close to zero if it is paid now, and high if it is paid later — retrofitting
means touching the lexer, the tags, the ribbon and the install docs at once.
So the rule is: **the shape of the code is general, the content is Python.**

### Where language-specific knowledge is allowed to live

Exactly one place: a `LangDef` record in a `modLang<Name>.bas` file. Nothing
else in the add-in may contain a Python-specific literal.

`modLangRegistry.bas` defines the `LangDef` type and holds the list. Adding a
language is:

1. Copy `modLangPython.bas`, fill in the table.
2. Add one `Register` line to `BuildRegistry`.

The lexer, renderer, palette, block code, gutter and ribbon XML are all
untouched by that.

### Why the rest of the design was already almost neutral

- **The palette is keyed by token class, not by syntax.** `Comment`, `String`,
  `Number`, `KeywordControl`, `KeywordDecl`, `Function`, `Class` are concepts
  every language in scope shares. Section 6 needed no changes at all.
- **The renderer only sees spans.** It has no idea what produced them.
- **The geometry spec in 5a is about font size and line count.** Nothing in it
  is Python-flavored.

### What did need changing

| Was | Now | Why |
|---|---|---|
| `PyCodeHighlight.ppam` | `CodeHighlight.ppam` | The filename ends up in install docs and on colleagues' machines. Renaming after distribution means a reinstall everywhere |
| Ribbon group "Python" | Groups "Code block" / "Highlight" / "Size", plus a Language dropdown | The dropdown is the visible extension point |
| Shape tags `PYCODE*` | `CODEBLOCK`, `CODEBLOCK_ID`, `CODEBLOCK_LANG`, `CODEBLOCK_GUTTER_OF` | Tags persist inside colleagues' decks. Changing them later orphans every block already in the wild |
| `modLexer` = Python tokenizer | `modLexer` = generic scanner + `modLangPython` table | The actual architectural change |

The tag rename and the file rename are the two that genuinely could not have
been deferred. Everything else could have been, and was still done now because
it was cheap.

### Deliberately not built yet

- **No second language table.** Writing a Java or C table now would be
  speculative — the table's field list is only proven correct once a second
  language actually uses it. The fields chosen are the ones Python needs plus
  block comments, which is the single most obvious gap.
- **No per-language palette overrides.** One theme, one token-class table.
- **No language auto-detection.** The block is tagged, and the tag is set from
  the dropdown. Guessing the language from the text is a Phase 6 nicety at best.

### The honest limit

`LangDef` is a table of *lexical* rules. It covers the C-like and Python-like
family comfortably — Java, C, C#, JavaScript, Go, Rust — because they all
tokenize as comments, quoted strings, numbers, identifiers and operators. It
will not stretch to a language whose lexing is structurally different: Lisp
readers, significant-whitespace string forms like Ruby heredocs, or anything
needing a preprocessor pass. If one of those is ever wanted, it gets its own
scanner and registers a different kind of entry. That is a fine outcome, and
worth stating so nobody tries to bend the table into a shape it cannot hold.

---

## 5a. LOCKED VISUAL SPEC

Chosen from the lab on 2026-08-14: **slide 5, Consolas 22pt, 1.2× spacing.**

Font size is the **only** free parameter. Everything else derives from it, so a
block can be resized on any slide without its proportions breaking. The VBA
renderer must use these formulas exactly — they are implemented and verified in
`tools/lab.py`.

```
size     S          user-adjustable, default 22pt
line     1.20 x S   EXACT points, never a spacing multiple
padding  0.64 x S   all four internal margins
gutter   digits x 0.550 x S  +  0.45 x S      (0.550 = Consolas advance/em)
radius   0.36 x S   corner radius, held visually constant
```

At the default 22pt that is: line 26.4pt, padding 14.1pt, gutter 34.1pt,
radius 7.9pt.

### Capacity per size

Measured against a 12.23in × 5.95in content area on a 16:9 slide.

| pt | lines | chars | |
|---|---|---|---|
| 10 | 34 | 154 | |
| 12 | 28 | 128 | |
| 14 | 24 | 109 | |
| 16 | 21 | 94 | legibility floor for a lecture hall |
| 18 | 18 | 83 | |
| 20 | 16 | 74 | |
| **22** | **15** | **67** | **default** |
| 24 | 13 | 61 | |
| 28 | 11 | 52 | |
| 32 | 10 | 44 | |

### Resizing rules

Three ribbon controls, all operating on the selected block:

- **A+ / A−** step through the ladder
  `10, 12, 14, 16, 18, 20, 22, 24, 28, 32`.
- **Fit** picks the largest ladder size at which the block's longest line and
  total line count both fit the content area. Paste code, press once, done.

Two guards, both needed because the lab showed what happens without them:

1. **A+ is capped at the fitting size.** At 32pt a 61-character line silently
   clipped off the right edge of the slide and the block overran the top. Growth
   must stop where the content stops fitting, not where the ladder ends.
2. **Warn below 16pt.** Auto-fit on a 27-line snippet correctly chose 12pt,
   which is arithmetically right and unreadable from the back of a room. When
   Fit lands below `MIN_TEACHING_SIZE = 16`, tell the user how many slides the
   snippet needs at 16pt instead of silently shrinking it. `split_advice()` in
   `tools/lab.py` already computes this.

Resizing must update, in one operation: font size on every run, exact line
spacing on every paragraph, all four margins, the gutter shape's width, font
size and margins, the block height, and the corner radius. Missing any one of
them is what makes a resized block look wrong.

---

## 6. The palette — `modTheme.bas`

VS Code Dark+, as RGB longs. Note that VBA's `RGB()` takes arguments in R, G, B
order but stores them as BGR internally — always build colors with `RGB()` and
never with a raw `&H` literal copied from a hex color, which will come out with
red and blue swapped.

| Token class | Hex | Used for |
|---|---|---|
| Background | `1E1E1E` | shape fill |
| Default | `D4D4D4` | operators, punctuation, plain identifiers |
| Comment | `6A9955` | `#` comments |
| String | `CE9178` | all string literals |
| Number | `B5CEA8` | numeric literals |
| Keyword control | `C586C0` | `if for while return try import from as` etc. |
| Keyword decl | `569CD6` | `def class lambda True False None and or not in is self` |
| Function | `DCDCAA` | definitions, call sites, builtins, decorators |
| Class / type | `4EC9B0` | after `class` |
| Gutter | `858585` | line numbers |

Keep every color behind a `ThemeColor(tokenClass)` function. Adding a light
theme later then means adding one table, not touching the renderer.

---

## 7. Rendering — `modRender.bas`

```
Sub ApplyHighlight(shp As Shape)
    tr = shp.TextFrame.TextRange
    tr.Font.Name = ThemeFontName()                ' set in Phase -1
    tr.Font.Color.RGB = ThemeColor(tkDefault)     ' reset everything first
    spans = Tokenize(tr.Text, GetLang(BlockLangId(shp)))
    For each span
        tr.Characters(span.Start, span.Length).Font.Color.RGB = ThemeColor(span.Class)
    Next
End Sub
```

Notes:

- **Reset first, then color.** This is what makes re-highlighting idempotent.
- **Merge adjacent same-class spans** in the lexer before rendering. Each
  `.Characters(...).Font.Color` assignment is a COM round trip and they dominate
  the runtime.
- **Skip default-colored spans** entirely — they were already handled by the
  reset.
- PowerPoint has no reliable `Application.ScreenUpdating`. If a 100-line block
  flickers badly, wrap the loop in the `LockWindowUpdate` Win32 API. Measure
  before adding it — for a 40-line snippet, roughly 300 spans, this should be
  well under a second.
- Do not touch bold or italic. Dark+ uses color alone, and leaving weight alone
  means the user can bold a line for emphasis and keep it.

---

## 8. Block creation and identity — `modBlock.bas`

`New code block` inserts a `msoShapeRoundedRectangle` and sets:

- Fill `1E1E1E`, no line, corner radius small
- `TextFrame2.WordWrap = msoFalse`, autosize none
- All four internal margins to a small fixed value, so gutter alignment is
  predictable
- Font Consolas, default 14pt, color `D4D4D4`
- Paragraph spacing zero, single line spacing
- Tags: `CODEBLOCK = "1"`, `CODEBLOCK_ID = <generated id>`,
  `CODEBLOCK_LANG = "python"`

**Identity strategy.** `Highlight` looks for the `CODEBLOCK` tag. If the
selection has no tag but is a shape with a text frame, it highlights it anyway
and adds the tag. That way pasting a code block from an old deck, or from a
colleague, still works. Nothing about display depends on the tag surviving.

**Language resolution**, in order: the shape's `CODEBLOCK_LANG` tag, then the
ribbon dropdown's current value, then `modLangRegistry.DefaultLangId()`. An
untagged block therefore highlights as Python without the user choosing
anything, which is the whole point of a default. An unrecognised tag value falls
back to the default rather than erroring — a block from a future version with
more languages should still render as *something*.

**Verify in Phase 1:** save the deck, close PowerPoint, reopen, confirm the tags
are still there. Shape tags are expected to persist in `.pptx`, but this is
worth ten minutes of proof before the design leans on it.

---

## 9. Line numbers — `modGutter.bas`

A separate text box to the left of the code block:

- Right-aligned, color `858585`, same font, same size, same line spacing
- Text is `"1" & vbCr & "2" & vbCr & ...`, one entry per paragraph in the code
  shape
- Tag `CODEBLOCK_GUTTER_OF = <the code block's CODEBLOCK_ID>`
- Positioned by setting `.Top` equal to the code shape's `.Top` and `.Left` to
  the code shape's `.Left` minus the gutter width

Rules:

- The toggle button creates the gutter if absent, deletes it if present.
- Every `Highlight` renumbers and re-aligns the existing gutter. So if the two
  shapes drift apart, pressing Highlight fixes it. That is the recovery path,
  and it means alignment does not need to be perfect on the first try.
- **Do not group the shapes.** Grouping makes text editing awkward and
  complicates every other operation. Accept that moving a block means moving two
  shapes, and that Highlight re-syncs them.
- Alignment depends on wrap being off and both shapes having identical font
  metrics and zero paragraph spacing. If a line wraps, alignment breaks by
  design.

---

## 10. Testing

VBA has no test framework, so this is visual inspection driven by a fixed corpus.

`modSelfTest.RunSelfTest` inserts every file in `tests/samples/<language>/` onto
its own scratch slide and highlights it with that language. Eyeball the result.
It takes a minute and catches regressions immediately.

Samples are filed per language — `tests/samples/python/` — so a second language
brings its own corpus rather than colliding in a shared directory. The failure
modes below are the Python list. Most of them have an obvious equivalent in any
other language, which is a decent sanity check on the `LangDef` fields.

Samples to write, each targeting one failure mode:

| File | Tests |
|---|---|
| `basic.py` | def, class, calls, builtins, plain assignment |
| `strings.py` | single, double, escaped quotes, `\\` at end of string |
| `triple.py` | docstrings containing `#`, a lone `"`, and the word `def` |
| `prefixes.py` | `r"" b"" f"" rb"" fr""`, uppercase prefixes |
| `numbers.py` | `0xFF 0b1010 1_000_000 1.5e-3 3j` |
| `comments.py` | `#` inside a string, string inside a comment, `#!` shebang |
| `decorators.py` | `@property`, `@app.route("/x")` |
| `long.py` | 150 lines, for timing and for gutter alignment |

---

## 11. Distribution

This is the part that will actually bite when you hand it to a colleague.

### Minimum version

**PowerPoint 2010 or newer, Windows desktop.** The binding constraint is
`customUI14.xml` ribbon markup. Everything else the design uses (`.ppam`,
`TextFrame2`, exact paragraph spacing, per-character font color, font
embedding) is 2007+.

PowerPoint for the web and the mobile apps have no VBA at all and can never run
this. They can still open and edit the decks it produces, which is the point of
the portability rules in section 2.

Verified on Sara's machine 2026-08-14 with `tools/check-env.ps1`:
Office Professional Plus 2021 Volume, build 16.0.14334.20806, 64-bit OS, macro
security at the Windows default. Comfortably above the floor.

`tools/check-env.ps1` is read-only and safe to send to the colleague ahead of
time. It reports version, build, edition, macro policy and Trusted Locations.

### Install steps

**Install into the default add-ins folder**, which is already a Trusted
Location on every stock Windows Office install:

```
%APPDATA%\Microsoft\Addins
```

Confirmed present in the Trusted Locations list on Sara's machine without any
configuration. Installing there means no macro prompt and no Trust Center work.

1. Copy `CodeHighlight.ppam` into `%APPDATA%\Microsoft\Addins`.
2. **Right-click the file > Properties > Unblock**, if it arrived by email,
   chat or download. A Trusted Location does bypass the mark-of-the-web block,
   but unblocking costs one click and removes a whole class of confusion.
3. PowerPoint > File > Options > Add-ins > Manage: **PowerPoint Add-ins** > Go >
   Add New > pick the file.
4. A "Code" tab appears in the ribbon.

**If a Trust Center policy still blocks it** — possible on an IT-managed
university machine, where `VBAWarnings` is set under the `Policies` registry key
and cannot be changed by the user — sign it. `SelfCert.exe` ships with Office
and creates a self-signed certificate. Sign the VBA project, and the colleague
trusts the publisher once on first load. That also survives updates to the file.
`check-env.ps1` reports whether the setting came from policy.

If their IT policy blocks add-ins outright, they still open and edit your decks
perfectly. They just cannot re-highlight. That is an acceptable failure mode and
was a design goal.

---

## 12. Build order

Roughly a day, plus the lab up front. Each phase ends in something you can
actually try.

**Phase −1 — visual lab (~30 min, do this first, separately).**
Generate `lab.pptx` from WSL, flip through the variants, pick the winner. Locks
down font, size, spacing, padding and palette. See section 2b. **Nothing below
starts until this is signed off**, because every later phase hardcodes these
numbers and changing them afterwards means touching the renderer, the block
creator and the gutter together.

**Phase 0 — packaging spike. DONE 2026-08-15.**
Empty `.ppam`, the real ribbon, every callback a `MsgBox`. Proved the customUI
injection and the install flow *before* writing any real code.

Result: Code tab loads, all seven controls fire, the Language dropdown populates
from `modLangRegistry`, every `imageMso` renders. Took about two hours rather
than thirty minutes, entirely because of the two silent-failure modes below.

**Both Phase 0 failures were silent, and that is the lesson.** Office does not
report a bad ribbon. It loads the add-in, loads the VBA, and shows no tab.

1. **Wrong relationship type.** `customUI14.xml` needs
   `.../office/2007/relationships/ui/extensibility`. The `2006` type belongs to
   the older `customUI.xml` part. The years in the namespace and the relationship
   type are offset by one, which is the entire trap. Diagnosed by checking
   `Application.AddIns` over COM: `registered` and `loaded` were both true, which
   ruled out the install and pointed straight at the part.
2. **Two `imageMso` ids that do not exist.** `Repaint` and `ListNumbering` both
   sound right and are not PowerPoint ids. A bad id renders blank and loads
   normally. `tools/check-imagemso.ps1` now tests ids against PowerPoint itself,
   and every id in the ribbon has been through it.

Font embedding was *not* settled here. It moves to Phase 4 — Consolas ships with
Windows, so it is belt-and-braces rather than load-bearing, and it was not worth
holding the MVP for.

Also settle font embedding here, since it is the other thing that can invalidate
the design: embed the chosen font in a test deck, open it on a machine without
that font installed, and confirm the block both renders and edits correctly.
Check whether `SaveAs ... EmbedTrueTypeFonts` is usable from VBA.

**Phase 1 — MVP. Lexer DONE 2026-08-15, renderer and block creation still to do.**
`New code block` plus a lexer covering comments, strings, numbers and keywords,
plus the renderer. At the end of this phase the add-in is genuinely usable.
Also: verify tag persistence across save and reopen.

The lexer is finished and verified: `tools/run-lexer-tests.sh` runs the scanner
over all ten samples and diffs it against `tools/lexref.py` character by
character. **All ten match.** The whole corpus tokenizes in 312 ms for 15,848
characters and 2,547 spans, so the performance worry in section 7 is settled -
no `LockWindowUpdate` needed.

Three real bugs the corpus caught, all now fixed, none of which would have been
obvious by eye:

- A decorator swallowed its dotted attribute, so `@name.setter` coloured
  `setter` as part of the decorator. Only the sigil and the first name belong
  to it.
- `cls(**kwargs)` came out blue. Being called outranks being a self word, the
  same way it outranks being a builtin.
- `match` was a hard keyword, so `re.match(...)` and any variable named `match`
  turned purple. It is a *soft* keyword - see `SoftKeywords` in the LangDef,
  which only fires where a statement can start.

Two defects in the reference classifier were also fixed, in `tools/lab.py`, so
the reference decks get them too. Both were cases where the VBA was right:
pygments splits `3j` into a number and a name, and it reads the `@` in `a @ b`
as a decorator. `tests/samples/python/decorators.py` was written specifically to
catch that second one.

**Phase 2 — identifier classification (~1 h).**
def / class / call site / builtins / self / decorators. This is what makes it
look like VS Code rather than like Notepad with green comments.

**Phase 3 — line numbers. DONE 2026-08-16.**
Gutter creation, toggle, renumber and re-align on every Highlight.

Alignment holds on the three properties the design already guaranteed: exact
point line spacing, wrap off, zero paragraph spacing. Measured on a 12-line
block - 12 numbers, matching top and left, code margin 48.2pt = pad 14.1 +
gutter 34.1, exactly the spec. Toggling off restores the margin.

Also built, unplanned, because Sara asked: **indent guides**, the vertical lines
marking nesting. Real line shapes in the indentation whitespace, so the text
stays pure source. Drift measured at 0.6pt or less on a 26.4pt line. Both the
gutter and the guides are per-block toggles - numbers off by default, guides on.

**Phase 4 — resizing and polish. DONE 2026-08-16.**
A+ / A− / Fit against the locked spec in section 5a, including the two guards
(cap A+ at the fitting size, warn below 16pt). Then Highlight all, and error
handling for the obvious wrong selections — nothing selected, a picture
selected, a group selected, an empty block.

Resizing touches every derived quantity at once, so build it against
`tools/lab.py --ladder` as the reference: render the same snippet at each ladder
size from VBA and compare against the Python output.

**Phase 5 — ship. PARTLY DONE. This is what is left, and it needs Sara.**

`INSTALL.md` is written. It covers install, the two autocorrect settings that
must be turned off, what every button does, the IT-policy fallback, and what a
colleague sees without the add-in.

Still to do, none of which can be done from WSL:

1. **Self-sign** with `SelfCert.exe`, so a machine with a policy-managed Trust
   Center can load it.
2. **Install on a second machine** following only what `INSTALL.md` says, and
   fix whatever the instructions get wrong.
3. **Open a deck on a machine with no add-in** and confirm it renders and edits.
   Consolas ships with Windows, so this should be uneventful - but it is the
   portability claim the whole design rests on, and it has never been tested.

**Phase 6 — optional, later.**
f-string interpolation. A light theme. A "strip highlighting" button to return a
block to plain text.

---

## 13. Known risks

| Risk | Mitigation |
|---|---|
| Triple-quoted string lexing is where this kind of scanner usually breaks | `triple.py` and `comments.py` samples exist specifically for this. Write them before the lexer, not after |
| Off-by-one from 1-based `Characters()` and CR paragraph marks | Test with a block whose first character is a keyword and whose last is a comment |
| Gutter misalignment | Wrap off, autofit off, zero paragraph spacing, and Highlight always re-aligns |
| Colleague's Trust Center blocks the add-in | Documented. Degrades to view-and-edit only, which is acceptable |
| Font embedding forgotten on a deck, so the block substitutes to a proportional font and alignment collapses | Add-in sets embedding automatically on block insert. Verified in Phase 0 on a machine without the font |
| Chosen font turns out not to be embeddable | Checked in Phase −1 before committing. Cascadia Mono is the fallback — system-wide on Windows 11, OFL, `.ttf` |
| Performance on very long blocks | Merge spans, skip defaults. If still slow, `LockWindowUpdate`. Set expectations around 150 lines per block |

---

## 14. Open questions for later

None blocking. Two worth revisiting once you have used it for a week:

- Should `Highlight` also normalize tabs to spaces? PowerPoint text boxes handle
  tab characters unpredictably, so probably yes, but it is a destructive edit to
  your source text and should be a deliberate choice.
- Should there be a keyboard shortcut for Highlight? PowerPoint does not let
  add-ins bind shortcuts directly, but the button can be pinned to the Quick
  Access Toolbar and then reached with Alt+number.
