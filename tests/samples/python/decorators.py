"""Decorators. The @ counts only as the first non-whitespace on a line."""

import functools


def memoize(fn):
    cache = {}

    @functools.wraps(fn)
    def wrapper(*args):
        if args not in cache:
            cache[args] = fn(*args)
        return cache[args]

    return wrapper


def tagged(label, *, strict=False):
    def decorate(fn):
        fn.label = label
        fn.strict = strict
        return fn
    return decorate


class Service:
    """Decorators indented inside a class body."""

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @staticmethod
    def version():
        return "1.0"

    @classmethod
    def build(cls, **kwargs):
        return cls(**kwargs)

    # A dotted decorator that takes arguments, the Flask-style case.
    @tagged("route", strict=True)
    def handler(self):
        return None


# Stacked decorators.
@memoize
@tagged("slow")
def fib(n):
    return n if n < 2 else fib(n - 1) + fib(n - 2)


# An @ that is NOT a decorator: the matrix-multiply operator, mid-expression.
class Matrix:
    def __matmul__(self, other):
        return "product"


a = Matrix()
b = Matrix()
product = a @ b
also_product = a@b

# An @ inside a string, and inside a comment.
email = "user@example.com"
handle = '@username'
# @not_a_decorator in a comment

print(fib(10), Service.version(), product, also_product, email, handle)
