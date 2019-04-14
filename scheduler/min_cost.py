#!/usr/bin/env python3
"""Contains min_cost, which returns a minimum-cost ordering of input values."""
def min_cost(teams, cost):
    """Returns a minimum-cost ordering of teams to slots with no more than one unpaired team.

    teams -- a list of teams to order; None indicates an empty slot
    cost -- the cost function used evaluate orderings; takes a list and return a number or tuple"""
    if len(teams) % 2: #this is designed for table pairs; that means even
        raise ValueError("Team list must contain an even number of elements")
    nones = sum(1 for team in teams if team is None)
    teams = [team for team in teams if team is not None]
    init_best = (cost(teams), teams + nones*[None])
    return _min_cost_internal(teams, cost, nones, [], init_best)[1]

def _min_cost_internal(teams, cost, nones, start_teams, best):
    """Basic branch-and-bound assignment sum solver; constructs orders to minimize lone teams."""
    start_cost = cost(start_teams)
    if not teams and start_cost < best[0]:
        best = (start_cost, start_teams + nones*[None])
    elif best[0] > start_cost and best[0] > cost([]):
        orders = []
        one_none = nones and not (len(start_teams) % 2 and start_teams[-1] is None)
        two_none = nones > 1 and len(start_teams) % 2 == 0
        for nones_used in [0] + one_none*[1] + two_none*[2]:
            orders += [([t for t in teams if t != team], nones - nones_used,
                        start_teams + nones_used*[None] + [team]) for team in teams]
        for new_teams, new_nones, new_start in sorted(orders, key=lambda order: cost(order[0])):
            best = min(best, _min_cost_internal(new_teams, cost, new_nones, new_start, best),
                       key=lambda x: x[0])
    return best
