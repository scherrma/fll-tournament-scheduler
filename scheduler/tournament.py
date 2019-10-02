#!/usr/bin/env python3
"""A module containing a Tournament class for using in creating FLL qualifier schedules."""
from datetime import timedelta, datetime
from numpy import gcd
import math
import scheduler.util as util
from scheduler.team import Team
import scheduler.min_cost

class Tournament:
    """A class designed to create schedules for FLL qualifier tournaments."""
    def __init__(self, teams, divisions, scheduling_method, travel, coach_meet, opening, lunch,
                 j_start, j_sets, j_calib, j_duration, j_break,
                 t_rounds, t_pairs, t_stagger, t_consec, t_duration):
        """Creates a tournament and requests a roster/settings file if one was not provided."""
        self.teams = [Team(*x) for x in teams]
        self.num_teams = len(self.teams)
        self.divisions = divisions
        self.scheduling_method = scheduling_method
        self.travel = travel
        self.coach_meet = coach_meet #first element is start time; second is duration
        self.opening = opening #first element is start time; second is duration
        self.lunch = lunch #first is earliest start time; second is latest; third is duration
        self.j_start = j_start
        self.j_sets = j_sets
        self.j_calib = j_calib
        self.j_duration = j_duration #first element is for judges; second is for teams
        self.j_break = j_break #first element is sessions between breaks; second is break duration

        self.t_rounds = t_rounds
        self.t_pairs = t_pairs
        self.t_stagger = t_stagger
        self.t_consec = t_consec
        self.t_duration = t_duration

        self.divs = []
        self.j_slots = []
        self.t_slots = []

    def schedule(self):
        """Top-level scheduling function; reads data and generates the schedule."""
        for team in self.teams:
            team.add_event(*self.coach_meet, 0, 0)
            team.add_event(*self.opening, 1, 0)

        print("Starting judge scheduling")
        self.split_divisions()
        if self.scheduling_method == "Interlaced":
            self.schedule_interlaced()
        elif self.scheduling_method == "Block":
            self.schedule_block()
        else:
            raise ValueError("{} scheduling is not supported".format(self.scheduling_method))
        self.assign_tables()

    def schedule_interlaced(self):
        """Top-level function controlling judge and table schedules for interlaced tournaments."""
        def run_rate():
            rate = 3*self.j_sets*self.t_duration[0]
            rate *= 1 + (self.t_consec < self.num_teams) / self.t_consec
            rate /= self.j_duration[0] + (self.j_break[1] / self.j_break[0] if
                    self.j_break[0] < self.num_teams else timedelta(0))
            return min(util.round_to(rate, 2), 2*self.t_pairs)

        matches_req = math.ceil((3*self.j_sets*self.j_break[0] - self.num_teams) / run_rate())
        time_req = matches_req * self.t_duration[0] + self.j_duration[1] + 2*self.travel
        time_ea = util.round_to(time_req / (self.j_break[0] - 1), timedelta(seconds=30))
        if timedelta(0) < time_ea - self.j_duration[0] <= timedelta(minutes=1):
            self.j_duration = (time_ea, self.j_duration[1])

        jlunch, jend = self.judge_interlaced_calib() if self.j_calib else self.judge_interlaced()

        print("Scheduling competition tables")

        time_increment = max(timedelta(minutes=1), gcd(self.t_duration[0], gcd(*self.j_duration)))
        offsets = [i*time_increment for i in range(1, self.t_duration[0] // time_increment)]
        offsets += [i*self.t_duration[0] for i in range(-3, 4)]
        offsets += [-x for x in offsets[1:]]

        earliest = max(sum(self.opening, 2*self.travel), self.j_slots[0][0])
        current, time_start = None, earliest
        best = ((datetime.max - self.travel, []), 0, earliest)
        while not current:
            current = min(((self.schedule_matches(time_start + offset, t, run_rate(),
                                                  range(self.t_rounds)[:2], True, jlunch, jend),
                                                  t, time_start + offset)
                           for t in range(self.num_teams) for offset in offsets
                           if time_start + offset >= earliest),
                          key=lambda x: (x[0][0], x[0][0] - x[2])) 
            if (current[0][0], current[0][0] - current[2]) < (best[0][0], best[0][0] - best[2]):
                best, current = current, False
                (end, self.t_slots), team_start, time_start = best 

        if self.t_rounds > 1: #determine run settings for afternoon table rounds
            for team in self.teams:
                team.add_event(team.next_avail(self.lunch[0], self.lunch[2]),
                               self.lunch[2], -1, -1)

            match_times = [(times[0], times[self.t_stagger] + self.t_duration[rnd])
                           for times, rnd, teams in self.t_slots
                           if any(team is not None for team in teams)]
            ref_lunch = max(match_times[i][0] - max(self.lunch[0], match_times[i-1][1])
                            for i in range(1, len(match_times))) > self.lunch[2]
            time_start = end + (max(self.t_duration[1:3]) if ref_lunch else self.lunch[2])

            current, end = None, datetime.max
            while not current:
                current = min((self.schedule_matches(time_start + self.t_duration[2] + offset,
                                                     (t_off + team_start) % self.num_teams,
                                                     None, range(self.t_rounds - 2))[0],
                              time_start + self.t_duration[2] + offset, t_off)
                              for offset in [tdelta for tdelta in offsets if tdelta >= timedelta(0)]
                              for t_off in range(math.ceil(self.num_teams / 2)))
                if current[:2] < (end, time_start):
                    end, time_start, t_offset = current
                    current = False
            team_start = (team_start + t_offset) % self.num_teams

            self.t_slots += [((self.t_slots[-1][0][0] + mdelta*self.t_duration[0],
                               self.t_slots[-1][0][1] + mdelta*self.t_duration[0]),
                              None, 2*self.t_pairs*[None]) for mdelta in
                             range(1, int((time_start - self.t_slots[-1][0][0]) / self.t_duration[0]))]
 
            self.t_slots += self.schedule_matches(time_start, team_start, None, range(2, self.t_rounds))[1]
            for team in self.teams:
                team.events = [event for event in team.events if event[2] != -1]

    def judge_interlaced(self): 
        """Generates the judging schedule for tournaments using interlaced scheduling.
        
           Does not work for tournaments with calibration rounds"""
        max_room = max(math.ceil(len(teams) / rooms) for rooms, teams in self.divs)
        self.divs = [(rooms, util.rpad(teams, util.round_to(len(teams), rooms), None))
                     for rooms, teams in self.divs]
        self.divs = [(rooms, sum(util.mpad(util.chunks(teams, rooms),
                                           max_room, rooms * [None]), []))
                     for rooms, teams in self.divs]

        max_room = max(math.ceil(len(div_teams) / rooms) for rooms, div_teams in self.divs)
        jrooms = [div_teams[i::rooms] for rooms, div_teams in self.divs for i in range(rooms)]
        jrooms = [util.rotate(jrooms[room], math.ceil(i * max_room / 3)) for i in range(3)
                  for room in range(self.j_sets)]

        self.teams = list({team : None for room in zip(*jrooms) for team in room if team is not None})
        team_dict = {team : i for i, team in enumerate(self.teams)}
        team_dict[None] = None

        jrooms = [[team_dict[team] for team in room] for room in jrooms]
        self.j_slots = [util.chunks(tslot, self.j_sets) for tslot in zip(*jrooms)]

        return self.assign_judge_times()

    def judge_interlaced_calib(self):
        """Generates the judging schedule for tournaments using interlaced scheduling.
        
           Does not work for tournaments with divisions"""
        teams = list(range(len(self.teams)))
        rot_dir = 1 + (len(teams) % 3 == 1)

        jslots = [sum([teams[(cat + j*rot_dir) % 3::3] for j in range(3)], [])
                  for cat in range(3)]
        jslots = [cat[:1] + (self.j_sets - 1)*[None] + cat[1:] for cat in jslots]
        self.j_slots = list(zip(*[util.chunks(cat, self.j_sets) for cat in jslots]))
        self.j_slots[-1] = [util.rpad(cat, self.j_sets, None) for cat in self.j_slots[-1]]
        self.j_slots[0] = ([0], [1], [2])

        return self.assign_judge_times()

    def schedule_block(self):
        """Generates judging and table schedules using block scheduling."""
        raise NotImplementedError("Block scheduling is not implemented yet")

    def split_divisions(self):
        """Sets self.divs to a list of (rooms for teams, teams) based on division."""
        max_room = max(12, math.ceil(self.num_teams / self.j_sets) + 1)
        rm_divs = [[team for team in self.teams if team.div == div]
                     for div in {team.div for team in self.teams}]
        rm_divs = [(math.ceil(len(div) / max_room), div) for div in rm_divs]
        
        total_room_req = sum(rooms for rooms, _ in rm_divs)
        if total_room_req > self.j_sets:
            self.divs, mixed_div = [], []
            teams_left, rooms_left = self.num_teams, self.j_sets
            for rooms_req, div_teams in rm_divs:
                if teams_left - len(div_teams) <= (rooms_left - rooms_req) * max_room:
                    self.divs.append((rooms_req, div_teams))
                    teams_left -= len(div_teams)
                    rooms_left -= rooms_req
                else:
                    mixed_div += div_teams
            if teams_left:
                self.divs.append((rooms_left, mixed_div))
        else:
            self.divs = rm_divs
            for i in range(self.j_sets - total_room_req):
                _, slow_div = max((len(teams) / rooms, idx)
                                  for idx, (rooms, teams) in enumerate(self.divs))
                rooms, teams = self.divs[slow_div]
                self.divs[slow_div] = (rooms + 1, teams)
        self.divs.sort(key=lambda x: sorted(list({team.div for team in x[1] if team})))
 
    def assign_judge_times(self):
        """Determines when each judging session will happen and assigns teams to those slots."""
        if self.j_break[0] > 1 and math.ceil((len(self.j_slots) - self.j_calib - 1) / self.j_break[0])\
                == math.ceil((len(self.j_slots) - self.j_calib - 1) / (self.j_break[0] - 1)):
                    self.j_break = (self.j_break[0] - 1, self.j_break[1])

        breaks = range(self.j_calib, len(self.j_slots) - 1, self.j_break[0])
        breaks = sorted(list({0, len(self.j_slots)} | set(breaks)))
        times = [[self.j_start + bool(i and self.j_calib)*self.travel
                  + max(i - self.j_calib, 0)*self.j_break[1] + j*self.j_duration[0]
                  for j in range(breaks[i], breaks[i+1] + 1)] for i in range(len(breaks) - 1)]
        j_blockers = [(start - self.travel, duration + 2*self.travel) for start, duration
                      in (self.opening, self.coach_meet)] + [self.lunch[1:]]

        for start, length in sorted(j_blockers):
            delay = max(timedelta(0), min(length, start + length - times[0][0]))
            delay -= self.j_break[1] if start >= times[0][-1] else timedelta(0)
            times = [[time + (start < cycle[-1])*delay for time in cycle] for cycle in times]
        times = [time for cycle in times for time in cycle]
        tdeltas = [(times[i + 1] - times[i], times[i] + self.j_duration[1]) for i in range(len(times) - 1)]
        lunch = max(tdeltas)[1] + self.j_duration[0] if max(tdeltas)[0] >= self.lunch[2] else None

        for breaktime in breaks[-2:0:-1]:
            self.j_slots.insert(breaktime, None)
        self.j_slots = list(zip(times, self.j_slots))
        for time, teams in filter(lambda x: x[1] is not None, self.j_slots):
            for cat, cat_teams in enumerate(teams):
                for room, team in filter(lambda x: x[1] is not None, enumerate(cat_teams)):
                    self.teams[team].add_event(time, self.j_duration[1], cat + 2, room)
        for i in range(len(self.j_slots) - 1, 0, -1):
            if self.j_slots[i][0] == self.j_slots[i - 1][0]:
                del self.j_slots[i - 1]
        return lunch, self.j_slots[-1][0] + self.j_duration[1]
    
    def schedule_matches(self, time_next, team_next, run_rate, rounds, lunch=False, jlunch=None, jend=None):
        time_finish, tslots = self.matches_inner(time_next, team_next, run_rate, rounds)
        if lunch and time_finish > self.lunch[1]:
            if jend is not None:
                lunch_time = (jlunch or jend) - self.j_duration[1]
            else:
                lunch_time = self.lunch[1] - self.travel
        
            for team in self.teams:
                team.add_event(lunch_time + self.travel, self.lunch[2], -1, None)
            time_finish, tslots = self.matches_inner(time_next, team_next, run_rate, rounds)
            for team in self.teams:
                team.events = [ev for ev in team.events if ev[2] != -1]

        return time_finish, tslots

    def matches_inner(self, time_next, team_next, run_rate, rounds):
        """Determines when table matches will occur and assigns teams to matches."""
        ideal_run_rate = 2*min(math.ceil((run_rate or 2*self.t_pairs)/2), self.t_pairs)

        def delay(t):
            return self._team(t + team_next).next_avail(time_next, window, self.travel) - time_next

        consec = 0
        last_nonnull, prev_nonnull = -1, -1
        tslots = []
        teams_left = len(rounds)*self.num_teams
        while teams_left > 0:
            rnd = rounds[len(rounds) - ((teams_left - 1) // self.num_teams + 1)]
            window = (1.5 if self.t_stagger else 1)*self.t_duration[rnd]
            run_rate = min(util.round_to(self.num_teams / math.ceil(self.travel / self.t_duration[rnd]
                                                                    + (3/2 if self.t_stagger else 1)), -2),
                           ideal_run_rate)
            match_sizes = (max(2, run_rate - 2), run_rate)

            max_teams, num_matches = next(filter(delay, range(teams_left)), teams_left), 0
            if max_teams:
                num_matches = math.floor((min(self._team(t + team_next).next_event(time_next)[0]
                                              for t in range(max_teams))
                                         - time_next - self.travel - int(self.t_stagger)*self.t_duration[rnd]/2) 
                                         / self.t_duration[rnd])
                num_matches = min(num_matches, math.ceil(delay(max_teams) / self.t_duration[rnd]
                                                         or teams_left / match_sizes[-1]))
            if max_teams < min(match_sizes) or num_matches == 0 or consec >= self.t_consec:
                consec = 0
                if num_matches and not teams_left <= max_teams <= match_sizes[-1]:
                    max_teams, num_matches = 0, 0

            next_matches = util.sum_to(match_sizes, max_teams, num_matches, max_teams >= teams_left)
            next_matches = next_matches[:self.t_consec - consec]
            next_matches = util.first_at_least(next_matches, (teams_left - 1) % self.num_teams + 1)

            for match_size in next_matches:
                timeslot = [(t + team_next) % self.num_teams for t in range(match_size)]
                tslots += [[(time_next, time_next + self.t_duration[rnd] / 2), rnd,
                                  util.rpad(timeslot, match_sizes[-1], None)]]
                time_next += self.t_duration[rnd]
                team_next, teams_left = team_next + match_size, teams_left - match_size
            consec += len(next_matches) if next_matches != [0] else 0

            if next_matches != [0]:
                prev_nonnull = len(tslots) - 2 if len(next_matches) > 1 else last_nonnull
                last_nonnull = len(tslots) - 1
            elif last_nonnull == 0 or (last_nonnull - prev_nonnull > 1):
                window = (1 + self.t_stagger/2)*self.t_duration[tslots[last_nonnull][1]]
                if all(self.teams[team].available(tslots[-1][0][0], window, self.travel)
                       for team in tslots[last_nonnull][2] if team is not None):
                    tslots[-1][1:], tslots[last_nonnull][1:] = tslots[last_nonnull][1:], tslots[-1][1:]
                    last_nonnull = len(tslots) - 1
                    consec = 1

        return time_next, tslots

    def assign_tables(self, assignment_passes=2):
        """Reorders the teams in self.t_slots to minimize table repetition for teams."""
        prev_tables = [[0 for i in range(2*self.t_pairs)] for j in range(self.num_teams)]
        def cost(order):
            val = sum(prev_tables[team][table]**1.1 for table, team in enumerate(order)
                      if team is not None)
            val += sum(self.t_rounds + 1 for i in range(0, len(order) - 1, 2) if
                       (order[i] is None) != (order[i + 1] is None))
            return val

        #the current approach only changes one match at a time; multiple passes fix bad early calls
        for assign_pass in range(assignment_passes):
            rotation = 0
            for(times, rnd, teams) in filter(None, self.t_slots):
                if assign_pass:
                    for table, team in filter(lambda x: x[1] is not None, enumerate(teams)):
                        prev_tables[team][table] -= 1
                else:
                    rotation += sum(2 for i in range(0, len(teams) - 1, 2)
                                    if teams[i] is teams[i + 1] is None)
                    rotation %= len(teams)
                teams[:] = scheduler.min_cost.min_cost(teams[rotation:] + teams[:rotation], cost)
                for table, team in filter(lambda x: x[1] is not None, enumerate(teams)):
                    prev_tables[team][table] += 1

        tbl_order = [2*j + k for i in range(2) for j in range(i, self.t_pairs, 2) for k in range(2)]
        for (times, rnd, teams) in filter(None, self.t_slots):
            teams[:] = util.rpad(teams, 2*self.t_pairs, None)
            teams[:] = [teams[tbl if self.t_stagger else i] for i, tbl in enumerate(tbl_order)]
            teams[:] = [(team, sum(event[2] > 4 for event in self._team(team).events)
                                   if team is not None else None) for team in teams]
            for table, (team, team_rnd) in filter(lambda x: x[1] != (None, None), enumerate(teams)):
                self._team(team).add_event(times[table >= util.round_to(self.t_pairs, 2)
                                                 and self.t_stagger],
                                           self.t_duration[rnd], 5 + team_rnd, table)
        self.clean_tslots()

    def clean_tslots(self):
        """Consolidates idle matches in self.t_slots."""
        def isnull(idx):
            return self.t_slots[idx] is None or all([team is None for pair in self.t_slots[idx][2] for team in pair])
        self.t_slots = self.t_slots[next(i for i in range(len(self.t_slots)) if not isnull(i)):]

        idx = 0
        while idx < len(self.t_slots):
            if isnull(idx) and self.t_slots[max(0, idx - 1)] is None:
                del self.t_slots[idx]
                idx = 0
            elif all([isnull(idx - i) for i in range(3)]):
                self.t_slots[max(0, idx - 2)] = None
                del self.t_slots[idx]
                del self.t_slots[max(0, idx - 1)]
                idx = 0
            else:
                idx += 1

    def _team(self, team_num):
        """Returns the team at the specified internal index; wraps modularly."""
        return self.teams[team_num % self.num_teams]
