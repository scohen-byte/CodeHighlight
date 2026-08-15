"""String prefixes. The prefix is part of the string token, and r disables escapes."""

import re

# Raw strings. The backslash does NOT escape the following character, so a
# raw string ending in a backslash is a syntax error in Python and must not
# appear here - but a backslash before a NON-quote is fine and common.
pattern = r"\d+\.\d+"
windows = r"C:\Users\User\ppt-lab"
raw_single = r'\n is two characters here'

# Bytes.
data = b"raw bytes \x00\xff"
data_single = b'more bytes'

# Unicode, legacy but still valid.
legacy = u"unicode literal"

# f-strings. Out of scope for v1 interpolation - the whole thing colors as one
# string, braces included.
name = "world"
greeting = f"hello {name}"
nested_quotes = f"outer {name!r} done"
expression = f"{1 + 2} and {len(name)}"

# Two-character prefixes, both orders.
raw_bytes = rb"\x00 raw"
bytes_raw = br"\x00 also raw"
raw_f = rf"\d {name}"
f_raw = fr"{name} \d"

# UPPERCASE and mixed-case prefixes are equally valid.
upper_raw = R"\d+"
upper_bytes = B"bytes"
upper_f = F"hello {name}"
mixed = Rb"\x00"

# A prefix letter immediately before a string but separated by an operator is
# NOT a prefix - this is the variable r, then a string.
r = 5
not_a_prefix = r + 1

# An identifier that merely ends in a prefix letter.
counter = 3
buffer = "value"

# Raw triple-quoted.
raw_triple = r"""
\d+ stays literal
"""

print(re.match(pattern, "1.5"), windows, raw_single, data, data_single, legacy)
print(greeting, nested_quotes, expression, raw_bytes, bytes_raw, raw_f, f_raw)
print(upper_raw, upper_bytes, upper_f, mixed, not_a_prefix, counter, buffer, raw_triple)
