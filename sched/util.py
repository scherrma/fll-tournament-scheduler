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

def positive(x):
    return x > 0

import datetime
def strtotime(in_val):
    """Converts a hh:mm formatted time string to a datetime."""
    if ':' not in in_val:
        h, m = int(in_val), 0
    elif in_val[-3] == ':':
        h, m = map(int, in_val.split(':'))
    else:
        raise ValueError
    return datetime.datetime(1, 1, 1, h, m)

def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in range(0, len(l), n):
        yield l[i:i + n]

def find_first(l, pred):
    return next((x for x in l if pred(x)))

def sum_to(options, goal, picks):
    select = []
    while goal > 0 and picks > 0:
        select += [min(options, key=lambda x:abs(x-goal/picks))]
        picks -= 1
        goal -= select[-1]
    return select if goal == picks == 0 else []
