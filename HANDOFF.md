# Handoff — updated 2026-08-15 (Sat)

`PLAN.md` holds the design and the full build order. This file holds only the
things `PLAN.md` does not: what physically exists right now, what has been
proven versus merely assumed, and the exact commands to pick the work back up.

**Phase −1 is complete and signed off. Phase 0 is half done — everything that
can be built from WSL exists and is tested. The remaining half needs PowerPoint
on Windows, and is the next thing to do.**

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
  ribbon/
    customUI14.xml            Code tab: 3 groups, 7 controls, 1 dropdown
  tools/
    lab.py                    visual lab AND the reference implementation of
                              the locked spec (derive / fit_size / split_advice)
    pack-ribbon.sh            injects the ribbon into a .ppam
    ppt.ps1                   PowerPoint COM helpers, called via powershell.exe
    check-env.ps1             read-only readiness check, safe to send to others
  tests/samples/python/
    lab_snippet.py            12 lines, the variant-comparison snippet
    long_snippet.py           27 lines, exercises auto-fit and splitting
  dist/                       generated, gitignored
    lab.pptx                  26 font/size/spacing variants
    ladder.pptx               locked spec at every ladder size, 10-32pt
```

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

## 4. NOT yet verified — check these early

These are the assumptions that could still invalidate part of the design.

1. **That PowerPoint accepts the packed `.ppam` and shows the tab.** The packer
   is proven to produce a structurally correct package. Whether Office agrees is
   the actual Phase 0 question and is still open.
2. **`imageMso` ids.** `SizeToControlHeight` on the Fit button is the least
   certain. A bad id renders blank rather than failing, so just look.
3. **Font embedding from VBA.** Whether `SaveAs ... EmbedTrueTypeFonts` is
   usable. Lower stakes than it was, because Consolas is on every Windows
   machine — embedding is now belt-and-braces rather than load-bearing.
4. **Shape tag persistence** across save, close and reopen. There is already a
   fallback if tags do not survive (treat any selected text box as code), so
   this is a convenience risk, not a design risk.
5. **Per-character coloring speed** in VBA at ~300 spans.

---

## 5. Resuming — exact commands

Everything runs from `/home/sara/projects/work/addin` in WSL.

```bash
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

## 6. Next move — finish Phase 0

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
