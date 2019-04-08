#  !/usr/bin/env python3
"""Utility functions for the fll tournament scheduler."""
import openpyxl
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

def basic_ws_format(ws, start=0):
    """bolds the top row, stripes the rows, and sets column widths"""
    for col in ws.columns:
        col[0].font = openpyxl.styles.Font(bold=True)
        length = 1.2*max(len(str(cell.value)) for cell in col[start:])
        ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = length
    for row in list(ws.rows):
        row[0].alignment = openpyxl.styles.Alignment(horizontal='right')
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            if cell.row > 1 and cell.row % 2 == 0:
                cell.fill = openpyxl.styles.PatternFill('solid', fgColor='DDDDDD')

def ws_borders(ws, borders=()):
    """Formats worksheet to add specified borders at a given period and offset."""
    for row in ws.rows:
        for cell in row:
            for border, period, offset, start_row in borders[::-1]:
                if (cell.column - offset) % period == 0 and cell.row > start_row:
                    cell.border = border
