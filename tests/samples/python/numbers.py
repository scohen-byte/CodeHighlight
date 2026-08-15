"""Numeric literals, and the identifiers that look like them."""

# Plain integers.
zero = 0
small = 7
big = 1234567

# Underscore separators, legal anywhere between digits.
million = 1_000_000
grouped = 0xFF_FF
binary_grouped = 0b1010_1010

# Floats, including the forms with an implicit zero.
half = 0.5
leading_dot = .5
trailing_dot = 5.
precise = 3.14159

# Exponents, upper and lower case, signed and unsigned.
small_exp = 1.5e-3
big_exp = 2.5E+10
bare_exp = 1e6
neg_exp = 1E-6

# Non-decimal bases, upper and lower case markers.
hex_lower = 0xff
hex_upper = 0XFF
octal = 0o755
octal_upper = 0O755
binary = 0b1010
binary_upper = 0B1010

# Complex numbers, both suffix cases.
imaginary = 3j
imaginary_upper = 3J
complex_float = 1.5j

# A number immediately followed by an operator, with no space.
tight = 1+2
ranged = range(0,10)

# Attribute access on a float literal.
rounded = (1.5).is_integer()

# Identifiers that START with or CONTAIN digit-like text but are not numbers.
x1 = 1
b1010 = "not binary"
o755 = "not octal"
xff = "not hex"
e10 = "not an exponent"
_1_000 = "leading underscore is an identifier"

# A number inside a string is a string.
version = "3.11.0"
hexish = "0xDEADBEEF"

# A number in a slice and a subscript.
items = [0, 1, 2, 3, 4]
sliced = items[1:3]
stepped = items[::2]

print(zero, small, big, million, grouped, binary_grouped)
print(half, leading_dot, trailing_dot, precise, small_exp, big_exp, bare_exp, neg_exp)
print(hex_lower, hex_upper, octal, octal_upper, binary, binary_upper)
print(imaginary, imaginary_upper, complex_float, tight, ranged, rounded)
print(x1, b1010, o755, xff, e10, _1_000, version, hexish, sliced, stepped)
