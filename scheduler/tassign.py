from math import isclose
from timeit import default_timer as timer

def tassign(slots, cost_fxn, mincost=0):
    if any(len(teams) % 2 for teams in slots):
        raise ValueError("Each slot must have an even length")
    new_slots = [[[t for t in slot if t is not None], sum(1 for n in slot if n is None)] for slot in slots]
    new_slots = [slot for slot in new_slots if slot[0]]

    num_tables = max(len(teams) + nones for teams, nones in new_slots)

    best = [[t for t in teams] + nones*[None] for teams, nones in new_slots]
    best = [best, cost_fxn(best)]

    _tassign_inner(new_slots, cost_fxn, [[] for slot in new_slots], best, mincost, timer() + 2*num_tables)

    best = best[0]
    for idx, teams in enumerate(slots):
        if teams and all(team is None for team in teams):
            best[idx : idx] = [teams]
    return best

def _tassign_inner(slots, cost_fxn, current, best, mincost, end_time):
    if isclose(best[1], mincost) or timer() > end_time:
        return
    if sum(len(teams) for teams, _ in slots) == 0:
        best[:] = [current, cost_fxn(current)]
        return
    
    slot_idx = min((len(teams), nones, i) for i, (teams, nones) in enumerate(slots) if teams)[-1]
    slot = slots[slot_idx]
    teams, nones = slot

    possibles = []
    if slot[1] > 1 and len(current[slot_idx]) % 2 == 0:
        new_current = [[t for t in teams] for teams in current]
        new_current[slot_idx] += [None, None]
        new_cost = cost_fxn(new_current), -len(new_current[slot_idx])
        if new_cost[0] < best[1]:
            new_slots = [[[t for t in teams], nones] for teams, nones in slots]
            new_slots[slot_idx][1] -= 2
            possibles.append((new_slots, new_current, [None, None], new_cost))

    for team in (nones > 0 and nones % 2)*[None] + slot[0]:
        new_current = [[t for t in teams] for teams in current]
        new_current[slot_idx].append(team)

        new_slots = [[[t for t in teams], nones] for teams, nones in slots]
        if team is None:
            new_slots[slot_idx][1] -= 1
        else:
            new_slots[slot_idx][0] = [t for t in new_slots[slot_idx][0] if t != team]

        new_teams, new_nones = new_slots[slot_idx]
        if len(new_teams) == 1 and len(new_current[slot_idx]) % 2:
            new_current[slot_idx].append(new_teams[0])
            new_teams, new_slots[slot_idx][0] = [], []
        if new_nones and not new_teams:
            new_current[slot_idx] += new_nones*[None]
            new_nones, new_slots[slot_idx][1] = 0, 0

        new_cost = cost_fxn(new_current), -len(new_current[slot_idx])
        if new_cost[0] < best[1]:
            possibles.append((new_slots, new_current, team, new_cost))
    possibles.sort(key=lambda x: x[-1])

    for new_slots, new_current, _, new_cost in possibles:
        if new_cost[0] < best[1]:
            _tassign_inner(new_slots, cost_fxn, new_current, best, mincost, end_time)
