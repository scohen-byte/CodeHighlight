# Handoff — paused 2026-08-15 (Sat), resuming Sunday 2026-08-16

`PLAN.md` holds the design and the full build order. This file holds only the
things `PLAN.md` does not: what physically exists right now, what has been
proven versus merely assumed, and the exact commands to pick the work back up.

**Phases −1 and 0 are complete. The lexer is done and passes its whole corpus.
Sunday starts with a THIN SLICE of the renderer — see section 6 — rather than
the full Phase 1, because the three things that could still go wrong only show
up in a real shape on a real slide.**

---

## 1. What exists on disk

```
addin/
  PLAN.md                     design + build order. The main document
  HANDOFF.md                  this file
  .gitignore
  src/
    modRibbon.bas             every ribbon callback. Stubs except the language
                              dropdown, which is real
    modLangRegistry.bas       LangDef type, the registry, word-set helpers
    modLangPython.bas         the Python table. Data only
    modTheme.bas              token classes and colours. Language-neutral
    modLexer.bas              the scanner. Names no language      <- the hard part
    modSelfTest.bas           runs the corpus, dumps masks for diffing
  ribbon/
    customUI14.xml            Code tab: 3 groups, 7 controls, 1 dropdown
  tools/
    lab.py                    visual lab AND the reference implementation of
                              the locked spec (derive / fit_size / split_advice)
    lexref.py                 reference classification as per-character masks
    run-lexer-tests.sh        the lexer test loop. Start here after a change
    lexer-test.ps1            its COM driver, called by the above
    pack-ribbon.sh            injects the ribbon into a .ppam
    check-imagemso.ps1        verifies ribbon icon ids against PowerPoint
    ppt.ps1                   PowerPoint COM helpers, called via powershell.exe
    check-env.ps1             read-only readiness check, safe to send to others
  tests/samples/python/
    lab_snippet.py            12 lines, the variant-comparison snippet
    long_snippet.py           27 lines, exercises auto-fit and splitting
    basic.py strings.py triple.py prefixes.py numbers.py
    comments.py decorators.py long.py       one failure mode each
  dist/                       generated, gitignored
    lab.pptx                  26 font/size/spacing variants
    ladder.pptx               locked spec at every ladder size, 10-32pt
    lexer/                    expected vs actual masks from the last test run
```

Also installed, on the Windows side:
`%APPDATA%\Microsoft\Addins\CodeHighlight.ppam` — the working add-in.

Now a git repo, as of 2026-08-15.

---

## 2. Decisions locked (do not relitigate)

| | |
|---|---|
| Name | **CodeHighlight**, renamed from PyCodeHighlight on 2026-08-15 |
| Font | **Consolas** — VS Code's Windows default, so it matches what students see |
| Default size | **22pt**, 1.2× line spacing (26.4pt exact) |
| Look | VS Code Dark Modern, `#1F1F1F` background, gutter inside the block |
| Chosen from | `lab.pptx` slide 5 |
| Resizing | Ladder 10–32pt, everything derived from size. Full formulas in `PLAN.md` §5a |
| Languages | Python only for now, but nothing outside `modLang*.bas` may contain a Python-specific literal. `PLAN.md` §5b |

The palette and every derived formula are implemented in `tools/lab.py`. **That
file is the spec.** When the VBA disagrees with it, the VBA is wrong.

---

## 3. Verified by running it, not assumed

2026-08-14:

- `powershell.exe` is callable from WSL. PowerPoint COM works through it,
  version 16.0.
- Slides export to PNG via COM, so rendered output can be inspected directly
  without opening PowerPoint by hand. This is the fast iteration loop.
- `python-pptx` 1.0.2 and `pygments` 2.20.0 are installed.
- Font metrics read from the real `.ttf` files with fontTools 4.63.0.
  Consolas advance 0.550 em, x-height 0.490 em, fsType Editable.
- The derived-geometry spec holds visually from 10pt through 32pt.
- Auto-fit and the split calculation produce correct numbers across
  12–40 line snippets.
- Sara's PowerPoint is Office Professional Plus 2021 Volume, build
  16.0.14334.20806, 64-bit OS, macro security at the Windows default. The floor
  is 2010, so there is plenty of headroom.
- `%APPDATA%\Microsoft\Addins` is already a Trusted Location. Installing the
  `.ppam` there avoids the macro prompt entirely. Run `tools/check-env.ps1` on
  the colleague's machine to confirm the same holds there.

2026-08-15:

- `tools/pack-ribbon.sh` works against a synthetic package: it injects the part,
  writes the relationship, is safe to re-run (no duplicate relationships), keeps
  `vbaProject.bin` byte-identical, and refuses both a non-zip and a zip with no
  VBA project in it.
- `ribbon/customUI14.xml` is well-formed, and so is the `.rels` the packer
  produces.
- `lab.py` still regenerates both decks after the samples moved into
  `tests/samples/python/`.

Also 2026-08-15, the big one:

- **The scanner matches the reference on all ten samples, character for
  character.** `tools/run-lexer-tests.sh`. That includes the triple-quote,
  string-escape and prefix cases the plan calls out as where scanners break.
- Tokenizing the whole corpus takes 312 ms for 15,848 characters and 2,547
  spans, so the performance worry is settled.
- The Code tab loads, all seven controls fire, the Language dropdown is fed by
  the registry, and every `imageMso` renders.

## 4. NOT yet verified — check these early

1. **Per-character colouring speed.** Tokenizing is fast, but the renderer's
   COM round trips are the part that was never the lexer's problem. 2,547 spans
   across ten files, so a single realistic block is roughly 100–300. Measure
   before reaching for `LockWindowUpdate`.
2. **Font embedding from VBA.** Whether `SaveAs ... EmbedTrueTypeFonts` is
   usable. Lower stakes than it was, because Consolas is on every Windows
   machine — embedding is now belt-and-braces rather than load-bearing.
3. **Shape tag persistence** across save, close and reopen. There is already a
   fallback if tags do not survive (treat any selected text box as code), so
   this is a convenience risk, not a design risk.
4. **Repeated `Application.Run` calls from PowerShell.** After two or three
   calls against a freshly imported project, further ones fail with "Sub or
   function not defined" for procedures that are demonstrably present. Never
   root-caused. The harness works around it by doing the whole corpus in one
   call. Anything new driven over COM should keep the call count low.

---

## 5. Resuming — exact commands

Everything runs from `/home/sara/projects/work/addin` in WSL.

```bash
# THE MAIN LOOP: run the scanner over the corpus and diff it against the
# reference. No human step. Add sample names to run only those.
tools/run-lexer-tests.sh
tools/run-lexer-tests.sh triple strings

# see what the reference says a file should look like
python3 tools/lexref.py --pretty tests/samples/python/triple.py

# regenerate the variant deck (26 slides)
python3 tools/lab.py

# regenerate the locked spec at every size (11 slides)
python3 tools/lab.py --ladder

# stage to Windows, then render to PNG for inspection
cp dist/ladder.pptx /mnt/c/Users/User/ppt-lab/
powershell.exe -NoProfile -ExecutionPolicy Bypass \
  -File "$(wslpath -w tools/ppt.ps1)" -Action export \
  -Path 'C:\Users\User\ppt-lab\ladder.pptx' \
  -Out  'C:\Users\User\ppt-lab\ladder-png'

# inject the ribbon into the add-in (PowerPoint must be CLOSED first)
tools/pack-ribbon.sh "/mnt/c/Users/User/AppData/Roaming/Microsoft/Addins/CodeHighlight.ppam"
```

Staging directory is `C:\Users\User\ppt-lab\`, which is
`/mnt/c/Users/User/ppt-lab/` from WSL. Artifacts must live on the C: drive —
COM against `\\wsl$\...` paths is slower and occasionally unreliable.

`ppt.ps1` also supports `-Action runmacro -Macro <name>`, which is how the VBA
gets tested from a script once the `.ppam` exists.

---

## 6. Sunday's first move — the thin slice

**Goal: one code block on one slide, highlighted, exported to PNG. Nothing
else.** No gutter, no resizing, no error handling, no Highlight-all. About 45
minutes, and it exists to fail fast in the same spirit as Phase 0.

Why this shape rather than building all of Phase 1 first: the lexer is proven
against files, but three things are unverified and every one of them only
appears once real text is in a real shape. Each is cheap to find now and
expensive to find after three more modules are stacked on top.

1. **The off-by-one.** The scanner has only ever seen file text with LF endings.
   PowerPoint uses **CR** for paragraph marks, a soft line break is **VT**, and
   `TextRange.Characters(start, length)` is **1-based**. Get it wrong and every
   colour shifts one character across the whole block — bizarre-looking and easy
   to misdiagnose. `PLAN.md` §13 calls this out. The mask harness CANNOT catch
   it: it tests the scanner, not the bridge from scanner to shape.
   Test with a block whose first character is a keyword and whose last is a
   comment.
2. **Per-character colouring speed.** Tokenizing is settled at 312 ms for the
   whole corpus, but that was never the risk. Each
   `.Characters(...).Font.Color` assignment is a COM round trip, and a realistic
   block is 100–300 spans. Measure before reaching for `LockWindowUpdate`.
3. **The COM call-count hazard in section 4.** Driving render tests means more
   `Application.Run` calls, which is exactly where it bit. Keep the count low.

### What the slice needs

- `modBlock.bas`, minimal: insert a `msoShapeRoundedRectangle`, fill `1F1F1F`,
  no line, wrap off, autofit off, all four margins from the `derive()` formulas
  in `tools/lab.py`, Consolas 22pt, **every paragraph explicitly left-aligned**,
  exact-point line spacing. Alignment is not optional — autoshapes default to
  centred and it is invisible until you look at a render.
- `modRender.bas`, minimal: reset the whole range to the default colour, then
  apply the spans, skipping default-coloured ones. Reset-then-colour is what
  makes re-highlighting idempotent.
- One driver macro that does insert + highlight in a single call, for the
  reason in section 4.
- Export the slide to PNG with `tools/ppt.ps1 -Action export` and look at it.

### The checkpoint

The PNG is the deliverable, not the code. It gets compared against
`dist/ladder.pptx` slide 5 — the signed-off look. If it matches, the rest of
Phase 1 is padding out modules along a proven path and becomes predictable.
If it does not, better to know before the gutter and the size ladder exist.

### Then, and only then, the rest of Phase 1

1. Gutter, resizing, Highlight-all, error handling for the obvious wrong
   selections — nothing selected, a picture, a group, an empty block.
2. Wire everything into the ribbon callbacks, replacing the stubs.
3. Verify shape tags survive save, close and reopen.

### On estimates

Phase 0 was estimated at 30 minutes and took about two hours. Both overruns
came from silent failures — Office reporting nothing at all rather than
reporting a problem. The render path has the same character, so treat "one hour
of code" as the floor and 2–3 hours as the realistic figure for a working MVP.

## 6a. Superseded — how Phase 0 finished

The WSL half is done. The rest is the part that cannot be scripted, because the
VBA project binary cannot be authored from WSL.

1. In PowerPoint on Windows: new blank presentation, Alt+F11, import
   `src/modLangRegistry.bas`, `src/modLangPython.bas`, `src/modRibbon.bas`.
   Import order does not matter. Then Save As → **PowerPoint Add-in (.ppam)** →
   `%APPDATA%\Microsoft\Addins\CodeHighlight.ppam`.
2. Close PowerPoint completely. It holds the file open and will overwrite a
   repack behind your back.
3. From WSL, run `tools/pack-ribbon.sh` on that path (command above).
4. Reopen PowerPoint → File → Options → Add-ins → Manage: PowerPoint Add-ins →
   Go → Add New → pick the file.
5. Confirm: a **Code** tab appears, all seven controls show, the Language
   dropdown lists "Python", and clicking anything raises the stub MsgBox.

If the tab does not appear, nothing else in the plan matters — so prove it
before writing the lexer. Only after that does Phase 1 start.

When importing a revised module later: **remove the old one first**, then
File → Import File. Importing without removing creates `modRibbon1`, which is a
confusing way to lose an afternoon.
