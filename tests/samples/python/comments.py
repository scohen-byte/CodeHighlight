#!/usr/bin/env python3
# The shebang above is a comment, including the ! and the /.
"""Comments, and the many things that merely look like comments."""

# A plain comment.
value = 1  # A trailing comment after code.

# A hash INSIDE a string is not a comment. Everything after it on this line
# is still code, and the trailing part must colour correctly.
url = "https://example.com/#anchor"
after_hash_string = url + "?q=1"  # but this one IS a comment

colour = '#1F1F1F'
selector = "#main > .item"
formatted = f"# {value} not a comment"

# A string INSIDE a comment is not a string:  "unterminated
still_code = 2

# An unterminated single quote inside a comment: it's fine
also_still_code = 3

# A triple quote inside a comment: """ does not open anything
definitely_still_code = 4

# A comment containing a backslash at the end of the line \
not_a_continuation = 5

# Multiple hashes.
## doubled
#### heading style

#no space after the hash
#

# A hash as the very first character of a line, with code after the comment.
counts = {"#": 1, "##": 2}

# A comment character inside a triple-quoted string.
doc = """
# not a comment
still inside the string
"""


def documented():  # comment on a def line
    """Docstring, then a comment below it."""
    # An indented comment.
    return "#"  # returning a hash


class Commented:  # comment on a class line
    # A comment as the first thing in a class body.
    field = "#value"


print(value, url, after_hash_string, colour, selector, formatted)
print(still_code, also_still_code, definitely_still_code, not_a_continuation)
print(counts, doc, documented(), Commented.field)
