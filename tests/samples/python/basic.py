"""Ordinary code. If this one looks wrong, nothing else matters."""

import math
from collections import Counter


class Account:
    rate = 0.05

    def __init__(self, owner, balance=0):
        self.owner = owner
        self.balance = balance
        self._history = []

    def deposit(self, amount):
        if amount <= 0:
            raise ValueError("amount must be positive")
        self.balance += amount
        self._history.append(amount)
        return self.balance

    def summary(self):
        total = sum(self._history)
        count = len(self._history)
        return f"{self.owner}: {count} deposits, {total} total"


def compound(principal, years, rate=Account.rate):
    for _ in range(years):
        principal *= 1 + rate
    return round(principal, 2)


accounts = [Account(name) for name in ("ada", "grace", "alan")]
tally = Counter(a.owner for a in accounts)
print(math.floor(compound(1000, 10)), tally)
