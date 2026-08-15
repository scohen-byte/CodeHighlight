import random


def bubble_sort(values):
    """Sort a list in place, the slow but readable way."""
    items = list(values)
    n = len(items)
    for i in range(n):
        swapped = False
        for j in range(0, n - i - 1):
            if items[j] > items[j + 1]:
                items[j], items[j + 1] = items[j + 1], items[j]
                swapped = True
        # Already sorted, so stop early.
        if not swapped:
            break
    return items


def check(size=12):
    data = [random.randint(0, 99) for _ in range(size)]
    result = bubble_sort(data)
    assert result == sorted(data), "bubble_sort is wrong"
    print(f"sorted {size} values: {result}")


check()
