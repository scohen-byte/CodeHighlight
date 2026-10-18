"""Prototype of the trace the add-in would run.

Runs a program once under sys.settrace and records, at each line event,
the line about to run, the watched names' values, and stdout so far.
Emits JSON. The add-in version would emit a tab-separated form.
"""
import io, json, sys

ABSENT = object()

def main(path, names):
    src = open(path, encoding="utf-8").read()
    code = compile(src, path, "exec")
    steps = []
    out = io.StringIO()
    real_stdout = sys.stdout

    def snapshot(frame):
        vals = {}
        for name in names:
            if name in frame.f_locals:
                v = frame.f_locals[name]
            elif name in frame.f_globals:
                v = frame.f_globals[name]
            else:
                vals[name] = None
                continue
            r = repr(v)
            if len(r) > 60:
                r = r[:57] + "..."
            vals[name] = {"repr": r, "id": id(v)}
        return vals

    def local_trace(frame, event, arg):
        if frame.f_code.co_filename != path:
            return None
        if event == "line":
            steps.append({"line": frame.f_lineno,
                          "func": frame.f_code.co_name,
                          "vars": snapshot(frame),
                          "out": out.getvalue()})
        elif event == "return" and frame.f_back is not None \
                and frame.f_back.f_code.co_filename == path:
            # The function is returning. Its locals are still alive here, and
            # the value it returns is known, so this is the state "after" its
            # return line. Control is about to land back on the caller's line,
            # which is where the slide points.
            steps.append({"line": frame.f_back.f_lineno,
                          "func": frame.f_code.co_name,
                          "ret": repr(arg)[:60],
                          "vars": snapshot(frame),
                          "out": out.getvalue()})
        return local_trace

    def global_trace(frame, event, arg):
        if frame.f_code.co_filename != path:
            return None
        return local_trace

    g = {"__name__": "__main__", "__file__": path}
    sys.stdout = out
    sys.settrace(global_trace)
    try:
        exec(code, g)
    finally:
        sys.settrace(None)
        sys.stdout = real_stdout
    # End state: nothing marked, module globals.
    end = {}
    for name in names:
        if name in g:
            r = repr(g[name])
            end[name] = {"repr": r[:60], "id": id(g[name])}
        else:
            end[name] = None
    steps.append({"line": 0, "func": "<end>", "vars": end, "out": out.getvalue()})
    json.dump({"source": src, "names": names, "steps": steps}, sys.stdout, indent=1)

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2:])
