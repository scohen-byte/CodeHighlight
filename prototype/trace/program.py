def total(nums):
    s = 0
    for n in nums:
        s = s + n
    return s

xs = [3, 1, 4]
t = total(xs)
print(t)
