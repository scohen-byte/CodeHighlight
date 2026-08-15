"""Triple-quoted strings. This is where scanners of this kind usually break.

This docstring itself contains a # which is not a comment,
a lone " which does not close anything,
and the word def which is not a keyword here.

    def looks_like_code(self):   # even indented like it
        return "nope"

Nothing above this line is code.
"""

import sys

SQL = '''
SELECT *
  FROM users          -- not a Python comment
 WHERE name = "quoted"
   AND note = 'single quoted inside triple single'
'''


def documented(x):
    """A normal docstring.

    Contains a # hash, a "double quote", and 'single quotes'.
    Ends with a quote character: "
    """
    return x


class Documented:
    '''Triple SINGLE quoted class docstring.

    It mentions """ without opening a triple double string.
    '''

    def method(self):
        """One-liner."""
        return None


# A triple-quoted string that ends on the same line it starts.
oneline = """all on one line"""
also = '''same here'''

# Two consecutive quotes are NOT a triple quote - this is an empty string
# followed by a comment marker inside a second string.
empty_then = "" + "# not a comment"

# A triple-quoted string containing an escaped quote before its terminator.
escaped_end = """ends with an escaped quote \\"""

# Adjacent triple strings.
back_to_back = """one""" """two"""

print(sys.argv, SQL, documented(1), Documented(), oneline, also)
print(empty_then, escaped_end, back_to_back)
