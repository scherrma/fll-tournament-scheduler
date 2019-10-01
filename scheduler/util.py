#  !/usr/bin/env python3
import math
"""Utility functions for the fll tournament scheduler."""
def sum_to(options, goal, picks, force_take_all=False):
    """Returns numbers that sum as close to a goal value as possible.

    options -- the numbers to select from
    goal -- the value to attempt to sum to
    picks -- the length of the list to be returned"""
    selected = []
    if not picks:
        selected = [0]
    elif force_take_all and goal < min(options):
        selected = [goal]
    else:
        while goal >= min(options) and picks > 0:
            selected.append(min((x for x in options if x <= goal),
                                key=lambda x: abs(x - goal/picks - 0.1)))
            goal -= selected[-1]
            picks -= 1
        if force_take_all and max(options) - selected[-1] >= goal:
            selected[-1] += goal
    return selected

def nth_occurence(ls, val, n):
    """Returns the index of the nth occurance of a value in a list."""
    return [i for i, x in enumerate(ls) if x == val][n - 1]

def round_to(val, base):
    """Returns the least integer multiple of base greater than val."""
    return math.ceil(val / base) * base

def rpad(ls, size, val):
    """Right-pads a list with a prescribed value to a set length."""
    return ls + (size - len(ls))*[val]

def mpad(ls, size, val):
    if size > len(ls):
        step = 1 / (size / len(ls) - 1)
        for i in range(size - len(ls), 0, -1):
            ls.insert(int((i + 1/2) * step), val)
    return ls

def first_at_least(ls, minimum):
    """Returns the first N elements of a list that sum to at least minimum (in-order)."""
    idx, total = 0, 0
    while idx < len(ls) and total < minimum:
        total += ls[idx]
        idx += 1
    return ls[:idx]

def rotate(ls, n):
    n = n % len(ls)
    return ls[n : ] + ls[ : n]

def chunks(ls, n):
    return [ls[i : i + n] for i in range(0, len(ls), n)]
