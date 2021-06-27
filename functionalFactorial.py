# implementing a function to find factorials in python


def functional_factorial(value):
    if value == 1:
        return value
    else:
        return value * functional_factorial(value - 1)


print(functional_factorial(3))