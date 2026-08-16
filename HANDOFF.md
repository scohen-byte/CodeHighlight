# Handoff — updated 2026-08-16 (Sun)

`PLAN.md` holds the design and the build order. `INSTALL.md` is written for
someone who has never seen the project. This file holds only what those two do
not: what exists right now, what has been proven versus assumed, and how to pick
the work back up.

**Phases −1 through 4 are complete. The add-in is built, installed and working.
What remains is Phase 5, and it needs Sara — see section 6.**

---

## 1. What exists

```
addin/
  PLAN.md  INSTALL.md  HANDOFF.md
  src/                        the add-in, 11 modules
    modTheme.bas              token classes and colours
    modSpec.bas               the locked geometry, as formulas, plus the ladder
    modLangRegistry.bas       LangDef type and the language list
    modLangPython.bas         the Python table. Data only
    modLexer.bas              the generic scanner        <- the hard part
    modBlock.bas              create / find / resize a block
    modRender.bas             spans onto the shape
    modGutter.bas             line numbers
    modGuides.bas             indent guides
    modRibbon.bas             commands, and the ribbon callbacks over them
    modSelfTest.bas           test scaffolding. NOT shipped
  ribbon/customUI14.xml       Code tab: 3 groups, 9 controls
  tools/
    build-addin.sh            ONE COMMAND BUILD. --install to deploy
    check-vba.sh              static checks. Run before driving PowerPoint
    run-lexer-tests.sh        the scanner vs the reference, whole corpus
    run-slice.sh              render one sample to PNG
    lexref.py / lab.py        the reference classifier and the visual lab
    pack-ribbon.sh  check-imagemso.ps1  check-env.ps1  ppt.ps1  probe.ps1
  tests/samples/python/       13 samples, each aimed at one failure mode
  dist/CodeHighlight.ppam     the shippable artifact
```

Installed at `%APPDATA%\Microsoft\Addins\CodeHighlight.ppam`, with the previous
build kept beside it as `.ppam.prev`.

---

## 2. The loop

```bash
tools/check-vba.sh              # static checks. ALWAYS run this first
tools/run-lexer-tests.sh        # scanner vs reference, all 13 samples
tools/run-slice.sh lab_snippet  # render one sample, and diff the shape's text
tools/build-addin.sh --install  # rebuild and deploy. PowerPoint must be CLOSED
```

`check-vba.sh` first is not a nicety. A VBA compile error does not come back as
an error: PowerPoint opens the VBA editor with a modal dialog, every later COM
call blocks, and the leftover process keeps the scratch file locked so the NEXT
run fails too. From outside that is indistinguishable from a hang. When
something hangs, kill `POWERPNT` and delete `C:\Users\User\ppt-lab\*.pptm`
before retrying.

---

## 3. Decisions locked (do not relitigate)

| | |
|---|---|
| Name | **CodeHighlight** |
| Font | **Consolas**, VS Code's Windows default |
| Default size | **22pt**, line spacing 26.4pt set in EXACT points |
| Background | `#1F1F1F`, VS Code Dark Modern |
| Comment green | `#87C76B`, brightened from Dark+'s `6A9955` for projection |
| Variables | **white** `#D4D4D4`, not light blue |
| Builtin classes | **teal**, including `range`, `zip`, `enumerate` |
| `and or not in is` | **purple**, with the control keywords |
| Languages | Python only, but nothing outside `modLang*.bas` may hold a Python-specific literal |

**The palette answers to the EDITOR, not to `lab.py`.** The reference uses
pygments, a plain grammar with no semantic model, and it was wrong in both
directions: it did not know `list` is a class, and it marked variables that VS
Code leaves alone. Every palette question in this project was settled by
counting pixels in a screenshot, and none by looking at a render.

---

## 4. Proven by measurement

- Scanner matches the reference on all 13 samples, character for character.
- Colouring costs ~0.7 ms per span. A 300-span block lands near 250 ms.
- Block geometry lands on the spec to the decimal: 12 lines at 22pt gives
  345.0pt predicted, 345.0pt measured.
- Shape tags survive save, close and reopen.
- Gutter alignment: 12 numbers for 12 lines, matching top and left, code margin
  48.2pt = pad 14.1 + gutter 34.1.
- Indent guides drift 0.6pt or less on a 26.4pt line.
- Size ladder agrees with the Phase −1 capacity table: 12 lines x 61 columns
  caps at 24pt. Larger is capped; Fit refuses to shrink below 16pt and reports
  how many slides are needed instead.
- Contrast against the block background: every token class clears WCAG AAA for
  large text. Comments were weakest at 4.95:1, now 8.17:1.

## 5. NOT verified

1. **A deck opened on a machine with no add-in.** The portability claim the
   whole design rests on, and it has never actually been tested.
2. **Font embedding from VBA.** Never settled. Low stakes - Consolas ships with
   Windows - but `SaveAs ... EmbedTrueTypeFonts` remains unexplored.
3. **A colleague's Trust Center.** Unsigned so far.
4. **Repeated `Application.Run` calls** from PowerShell go wrong after a few
   calls against a freshly imported project, in a way never root-caused. The
   harnesses work around it by doing whole runs in one macro call. Keep the call
   count low in anything new.
5. **The builtin classes coloured teal were INFERRED, not measured.** Only
   `list` was verified against a screenshot. `range`, `zip`, `enumerate`,
   `map`, `filter`, `reversed` and `super` are classes in CPython and are
   coloured as such - but if VS Code shows any of them yellow, moving them from
   `TypeNames` back to `Builtins` in `modLangPython.bas` is a one-line fix.

---

## 6. What is left — Phase 5, and it needs Sara

`INSTALL.md` is written and covers install, use, the two autocorrect settings,
and the IT-policy fallback. What remains cannot be done from WSL:

1. **Open a deck on a machine without the add-in.** This is the most important
   one. Everything about the design - no custom XML parts, no external files,
   plain shapes with per-character colours - exists to make this work, and it
   has never been checked. Do this before signing anything.
2. **Self-sign with `SelfCert.exe`**, so a policy-managed Trust Center can load
   it. `check-env.ps1` reports whether a machine's setting came from policy.
3. **Install on a second machine following only `INSTALL.md`**, and fix whatever
   the instructions get wrong. Written instructions are always wrong the first
   time they are used by someone else.

### Smaller things, whenever

- **Guides and gutter are per-block toggles.** Guides default on, numbers off.
- **A "strip highlighting" button**, to return a block to plain text. In
  `PLAN.md` §12 as a Phase 6 idea and still unbuilt.
- **f-string interpolation is done**, contrary to what §5 originally said.
- **Orphaned guides.** Delete a block and its guide lines stay on the slide,
  since nothing is left to tie them to. Moving a block is fine — Highlight
  re-syncs. Deleting is not handled.
- **Tabs are not converted to spaces.** Tab stops are set to 4 characters so
  they LOOK right, but the characters stay tabs. `PLAN.md` §14 wanted that
  choice made deliberately, and it has not been made.
