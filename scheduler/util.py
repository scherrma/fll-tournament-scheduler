#!/usr/bin/env python3
"""Utility functions for the fll tournament scheduler."""
def sum_to(options, goal, picks):
    """Returns numbers that sum as close to a goal value as possible.

    options -- the numbers to select from
    goal -- the value to attempt to sum to
    picks -- the length of the list to be returned"""
    select = []
    if goal < min(options):
        select = [goal]
    else:
        while goal >= min(options) and picks > 0:
            select += [min(options, key=lambda x: abs(x-goal/picks))]
            picks -= 1
            goal -= select[-1]
    return select

def nth_occurence(ls, val, n):
    """Returns the index of the nth occurance of a value in a list."""
    return [i for i, x in enumerate(ls) if x == val][n - 1]

def rpad(ls, size, val):
    """Right-pads a list with a prescribed value to a set length."""
    return ls + (size - len(ls))*[val]
