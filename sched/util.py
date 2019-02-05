#!/usr/bin/python3
def clean_input(question, validate = None, parse = None):
    """Continually requests input from terminal until given something parseable that passes validate"""
    while True:
        answer = input(question).strip()
        try:
            value = parse(answer) if parse is not None else answer
            if validate is None or validate(value):
                return value
        except ValueError:
            pass
        print(answer, "is not a valid input")    

def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in range(0, len(l), n):
        yield l[i:i + n]

def find_first(l, pred):
    return next((x for x in l if pred(x)))

def floor_to(num, base):
    return num - (num % base)

def sum_to(options, goal, picks):
    select = []
    while goal > 0 and picks > 0:
        select += [min(options, key=lambda x:abs(x-goal/picks))]
        picks -= 1
        goal -= select[-1]
    return select
