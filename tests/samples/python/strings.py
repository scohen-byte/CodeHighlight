"""Single-line string edge cases.

Every line here is a place a naive scanner stops in the wrong column.
"""

plain = 'single quoted'
double = "double quoted"

# A quote of the other kind inside a string is just a character.
apostrophe = "it's fine"
quotation = 'she said "hello" and left'

# Escaped quotes of the SAME kind. The escape must consume the quote so the
# string does not end early.
escaped_single = 'don\'t stop here'
escaped_double = "say \"hello\" properly"

# A backslash at the very end of the content, immediately before the closing
# quote. The escape consumes the second backslash, NOT the quote.
trailing_backslash = "ends with a backslash \\"
also_trailing = 'and so does this \\'

# Two escapes in a row, then a real quote.
doubled = "a\\\\b"

# An escape sequence that is not a quote at all.
newline = "line one\nline two"
tabbed = "col\tcol"

# Empty strings, including an empty one immediately followed by another string.
empty = ""
empty_pair = "" + ''
adjacent = "back" "to" "back"

# A string containing what looks like the start of a triple quote.
not_triple = "this has \"\" two quotes"

# Concatenation across a line continuation.
continued = "first part " \
            "second part"

print(plain, double, apostrophe, quotation, escaped_single, escaped_double)
print(trailing_backslash, also_trailing, doubled, newline, tabbed)
print(empty, empty_pair, adjacent, not_triple, continued)
