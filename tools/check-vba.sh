#!/usr/bin/env bash
#
# Static checks on the .bas sources, for the mistakes VBA reports badly.
#
#   tools/check-vba.sh
#
# Run this before driving PowerPoint. A compile error does not come back as an
# error from Application.Run - it opens the VBA editor with a modal dialog, and
# every subsequent COM call blocks. From the outside that is indistinguishable
# from a hang, and it leaves a PowerPoint process holding the scratch file, so
# the NEXT run fails too. One of these cost a long detour on 2026-08-16.

set -uo pipefail
REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO/src" || exit 1
fail=0

for f in *.bas; do
    # 1. Balanced procedure declarations.
    d=$(grep -cE '^[[:space:]]*(Public |Private )?(Sub|Function) ' "$f")
    e=$(grep -cE '^[[:space:]]*End (Sub|Function)$' "$f")
    if [[ "$d" != "$e" ]]; then
        echo "$f: $d procedure declarations but $e End statements"; fail=1
    fi

    # 2. Ambiguous names within a module.
    dup=$(grep -oE '^[[:space:]]*(Public |Private )?(Sub|Function) [A-Za-z_]+' "$f" \
          | awk '{print $NF}' | sort | uniq -d)
    if [[ -n "$dup" ]]; then
        echo "$f: duplicate procedure name(s): $dup"; fail=1
    fi

    # 3. Module-level declarations must ALL precede the first procedure. VBA
    #    reports a violation as "Variable not defined" at the USE site, which
    #    sends you looking in entirely the wrong place.
    fp=$(grep -nE '^(Public |Private )?(Sub|Function|Property) ' "$f" | head -1 | cut -d: -f1)
    ld=$(grep -nE '^(Public|Private) (Const |Type |Enum |[A-Za-z_]+ As |[A-Za-z_]+\(\) As )' "$f" \
         | tail -1 | cut -d: -f1)
    if [[ -n "$fp" && -n "$ld" && "$ld" -gt "$fp" ]]; then
        echo "$f: module declaration on line $ld comes after the first procedure on line $fp"; fail=1
    fi
done

# 4. Duplicate PUBLIC names across modules, which is also ambiguous.
dup=$(grep -hoE '^Public (Sub|Function) [A-Za-z_]+' *.bas | awk '{print $3}' | sort | uniq -d)
if [[ -n "$dup" ]]; then
    echo "duplicate public procedure name(s) across modules: $dup"; fail=1
fi

# 5. Every ribbon callback named in the XML must exist in modRibbon. A missing
#    one does not fail at load: the tab appears, and the button reports
#    "Cannot run the macro" when clicked - which reads like a packaging fault.
RIBBON="$REPO/ribbon/customUI14.xml"
if [[ -f "$RIBBON" ]]; then
    for cb in $(grep -oE '(onAction|getText|getPressed|getItemCount|getItemLabel|getSelectedItemIndex|onChange|onLoad)="[A-Za-z0-9_]+"' "$RIBBON" \
                | cut -d'"' -f2 | sort -u); do
        if ! grep -q "Sub $cb(" modRibbon.bas; then
            echo "ribbon callback '$cb' is named in customUI14.xml but not defined in modRibbon.bas"; fail=1
        fi
    done
fi

if [[ $fail -eq 0 ]]; then
    echo "vba checks passed ($(ls *.bas | wc -l) modules, $(grep -coE '(onAction|getText|getPressed|getItemCount|getItemLabel|getSelectedItemIndex|onChange|onLoad)="' "$RIBBON" 2>/dev/null || echo 0) ribbon callbacks)"
fi
exit $fail
