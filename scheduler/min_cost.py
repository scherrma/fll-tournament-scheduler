#!/usr/bin/env python3
"""Contains min_cost, which returns a minimum-cost ordering of input values."""
def min_cost(teams, cost):
    """Returns a minimum-cost ordering of teams to slots with no more than one unpaired team.

    teams -- a list of teams to order; None indicates an empty slot
    cost -- the cost function used evaluate orderings; takes a list and return a number or tuple"""
    if len(teams) % 2:
        raise ValueError("Team list must contain an even number of elements")
    nones = sum(1 for team in teams if team is None)
    teams = [team for team in teams if team is not None]
    return _min_cost_internal(teams, cost, nones, [], (cost(teams), teams))[1]

def _min_cost_internal(teams, cost, nones, start_teams, best):
    start_cost = cost(start_teams)
    if not teams:
        best = min(best, (start_cost, start_teams + nones*[None]), key=lambda x: x[0])
    elif best[0] > start_cost and best[0] > cost([]):
        none_allowed = (nones > 1 and len(start_teams) % 2 == 0) or\
                       (nones % 2 and len(start_teams) % 2)
        for ending in teams + none_allowed*[None]:
            new_teams, new_start, new_nones = teams[:], start_teams[:], nones
            if ending is not None:
                new_teams = [team for team in teams if team != ending]
                new_start = start_teams + [ending]
            else:
                new_nones = nones + len(start_teams) % 2 - 2
                new_start = start_teams + (nones - new_nones)*[None]

            best = min(best, _min_cost_internal(new_teams, cost, new_nones, new_start, best),
                       key=lambda x: x[0])
    return best
