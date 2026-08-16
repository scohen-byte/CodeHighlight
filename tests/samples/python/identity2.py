import copy
x = [1,2,3]
y = copy.copy(x)
z = list(x)

print(y is x) #False
print(z is x) #False
