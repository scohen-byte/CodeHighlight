# Trace: a walkthrough driven by running the program

Working notes for a feature that does not exist yet. Started 2026-10-18.
Nothing in here is wired into the add-in. `src/` is untouched.

## The idea

A fourth walkthrough button. Select a variable name in a block, press
**Trace**, and the program is run once under a tracer. One slide is made per
step of the run, each marking the line and showing the watched variables'
values at that point, and optionally what has been printed so far.

The run happens at authoring time, on the lecturer's own machine. The deck
that comes out is ordinary slides with ordinary text. In class you advance.
That is what makes it feasible: nothing executes on the podium machine, the
ribbon's absence during a slideshow does not matter, Defender's rule against
Office spawning child processes does not matter, and the deck stays portable.

It is the existing `DoStepThrough` in `modRibbon.bas` with two changes: the
list of steps comes from the run rather than from the text, so a line can
repeat, and each step carries values to display.

## What is here

- `tracer.py` is the prototype of the script the add-in would run. It uses
  `sys.settrace`, records each line event with the watched names' values and
  stdout so far, and records each return event too. About 60 lines. In the
  add-in it would live as a string constant, written to a temp file, and would
  emit a tab-separated form rather than JSON, since VBA has no JSON parser.
- `program.py` is the sample: a function summing a list in a loop, called
  once, result printed. Nine lines, 13 line events and one return.
- `trace.json` is the tracer's output for it, watching `s`, `n`, `t`.
- `mockup.html` is the click-through mockup built from that trace. It is also
  published at https://claude.ai/artifact/QfqzpKeFUUjzmBjs4kHkDP so it can be
  opened anywhere. It renders the slides at the add-in's geometry and palette
  (modSpec, modTheme) with a switch for every open decision.

To regenerate the trace:

    python3 -I tracer.py program.py s n t > trace.json

## Decided

**Values are shown AFTER the line runs.** Tried both in the mockup. After
reads naturally. The one place it broke was `return`: the state after a
return line, taken from the next line event, is back in the caller, where
the function's locals no longer exist, so the slide pointing at `return s`
showed `s` as undefined. Fixed in the tracer by recording the return event as
a step of its own: it points at the caller's line, the returning frame's
locals are still alive, and the returned value is known. In the mockup that
reads as "total() returned 8" with `s = 8` still showing, and the next slide,
back at the call, shows `t = 8`.

## Still open, and what the mockup offers for each

1. **Where the values show.** Fixed panel beside the block, a note on the
   current line, or inline ghost text after the line. The panel can show the
   scope ("inside total()", "at top level") and band a value that changed.
2. **Which steps get a slide.** Every executed line, only selected lines, or
   only when a watched value changes. A loop over three lines for ten
   iterations is thirty steps, so this is also the answer to the twelve-slide
   guard. The change-only filter cut the sample from 15 slides to 9 with
   nothing lost.
3. **Arrow or band.** Point at is probably the right base: fading the rest of
   the code while a variable changes hides the loop being traced.

Sara judges this from the rendered page, not from a description.

## Known constraints

- The block must be a whole program. Fragments, placeholders and code split
  across slides with First line do not run.
- `input()` needs scripted answers. Randomness should be seeded so re-running
  gives the same slides.
- Values are reprs, truncated. An instance without its own repr shows an
  address. Aliasing wants identity, not value: two names holding the same
  object could be marked as such.
- A name is looked up in the running frame's locals, then globals, and shown
  as "not defined" when absent. A local that vanishes on return is correct
  and is the point of the lesson.
- Needs Python on the authoring machine and a timeout on the run. The
  environment check script should report both.
- Hebrew in output needs Python forced to UTF-8 and the capture read as UTF-8.
