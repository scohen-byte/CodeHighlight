# Handoff — updated 2026-08-15 (Sat)

`PLAN.md` holds the design and the full build order. This file holds only the
things `PLAN.md` does not: what physically exists right now, what has been
proven versus merely assumed, and the exact commands to pick the work back up.

**Phases −1 and 0 are complete. The lexer is done and passes its whole corpus.
Next: the renderer and block creation, which finish the Phase 1 MVP.**

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

## 6. Next move — finish the Phase 1 MVP

The lexer is done. `tools/run-lexer-tests.sh` runs the VBA scanner over all ten
samples and diffs it against `tools/lexref.py` character by character, and all
ten match. It needs no human step: it imports the modules into a scratch deck
over COM, runs them, and diffs. Use it after every scanner change.

Remaining for the MVP:

1. `modBlock.bas` — `New block` inserts a rounded rectangle against the locked
   spec in `PLAN.md` §5a. Left-align every paragraph explicitly, wrap off,
   autofit off, exact point line spacing.
2. `modRender.bas` — apply spans to the shape. Reset to the default colour
   first, then colour, which is what makes re-highlighting idempotent. Skip
   default-coloured spans; they were handled by the reset.
3. Wire both into the ribbon callbacks, replacing the stubs.
4. Verify shape tags survive save, close and reopen.

Performance is already settled: the whole corpus tokenizes in 312 ms for 15,848
characters and 2,547 spans, so `LockWindowUpdate` is not needed.

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
