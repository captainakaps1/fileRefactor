# implementing a function to find factorials in python


def tail_call_factorial(value, num=1):
    if value == 1:
        return num
    return tail_call_factorial(value - 1, value * num)


print(tail_call_factorial(3))
