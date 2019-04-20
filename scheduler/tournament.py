#!/usr/bin/env python3
"""A module containing a Tournament class for using in creating FLL qualifier schedules."""
from datetime import timedelta, datetime
import math
import scheduler.util as util
from scheduler.team import Team
import scheduler.min_cost

class Tournament:
    """A class designed to create schedules for FLL qualifier tournaments."""
    def __init__(self, teams, divisions, scheduling_method, travel, coach_meet, opening, j_start,
                 j_sets, j_calib, j_duration, j_break, j_lunch, t_rounds, t_pairs, t_duration, t_lunch):
        """Creates a tournament and requests a roster/settings file if one was not provided."""
        self.teams = [Team(*x) for x in teams]
        self.num_teams = len(self.teams)
        self.divisions = divisions
        self.scheduling_method = scheduling_method
        self.travel = travel
        self.coach_meet = coach_meet #first element is start time; second is duration
        self.opening = opening #first element is start time; second is duration
        self.j_start = j_start
        self.j_sets = j_sets
        self.j_calib = j_calib
        self.j_duration = j_duration #first element is for judges; second is for teams
        self.j_break = j_break #first element is sessions between breaks; second is break duration
        self.j_lunch = j_lunch #first element is start time; second is duration

        self.t_rounds = t_rounds
        self.t_pairs = t_pairs
        self.t_duration = t_duration
        self.t_lunch = t_lunch #first element is start time; second is duration

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
        print("Scheduling competition tables")
        self.assign_tables()

    def schedule_interlaced(self):
        """Top-level function controlling judge and table schedules for interlaced tournaments."""
        self.judge_interlaced()
        print("Starting table scheduling")

        #determine how morning table rounds will operate, then schedule them
        time_start = [e_start + e_length + self.travel for (e_start, e_length, *_)
                      in self.teams[3*self.j_calib].events]
        time_start = sorted([t for t in time_start if t >= self.j_start])
        time_start = next((time for time in time_start if self.teams[3*self.j_calib]
                           .available(time, self.t_duration[0], self.travel)))
        team_idx = next((t for t in range(3*self.j_calib + 1) if
                         self._team(t).available(time_start, self.t_duration[0], self.travel)))

        run_rate = 3*self.j_sets*self.t_duration[0]/(self.j_duration[0] + self.j_break[1]/self.j_break[0])

        min_early_run_rate = max(2, 2*math.ceil(run_rate/2))
        if self.num_teams % min_early_run_rate:
            early_avail = [self._team(team_idx - i).available(time_start - self.t_duration[0],
                                                              self.t_duration[0], self.travel)
                           for i in range(min_early_run_rate)]
            if all(early_avail):
                team_idx -= len(early_avail)
                time_start -= self.t_duration[0]

        current_end = datetime.max
        backup_end = current_end
        time_start -= self.t_duration[0]
        while current_end <= backup_end:
            backup_tslot, self.t_slots = self.t_slots, []
            backup_end = current_end

            time_start += self.t_duration[0]
            current_end = self.schedule_matches(time_start, team_idx, range(min(2, self.t_rounds)),
                                                run_rate)
        else:
            time_start -= self.t_duration[0]
            self.t_slots = backup_tslot

        if self.t_rounds > 1: #determine run settings for afternoon table rounds
            self.t_slots += [None]
            time_restart = [sum(self._team(t).events[-1][:2], self.travel
                                - t // (2*self.t_pairs) * self.t_duration[2])
                            for t in range(self.num_teams)]
            time_start = max(time_restart + [backup_end + self.t_lunch[1]])
            self.schedule_matches(time_start, 0, range(2, self.t_rounds), None)

    def judge_interlaced(self):
        """Generates the judging schedule for tournaments using interlaced scheduling."""
        self.j_calib = self.j_calib and not self.divisions

        max_room = max([math.ceil(len(teams) / rooms) for teams, rooms in self.divs])
        most_rooms = max([rooms for teams, rooms in self.divs])

        #teams are selected three times (once for each category) in monotonically increaseing order
        #so team 2 will first see judges no later than team 1's second trip but no earlier than
        #team 1's first trip
        picks = [util.rpad(rooms*[i], most_rooms, None) for i, (tms, rooms) in enumerate(self.divs)]
        pick_order = math.ceil(max_room / 3)*[val for div_picks in zip(*picks)
                                              for val in 3*div_picks if val is not None]

        for div, (teams, rooms) in enumerate(self.divs):
            last_team = util.nth_occurence(pick_order, div, len(teams))
            pick_order = [x for idx, x in enumerate(pick_order) if idx <= last_team or x != div]

        #to ensure teams stay in order we sometimes idle judging rooms
        excess = min([-len(teams) % 3 for teams, rooms in self.divs if
                      math.ceil(len(teams) / rooms) == max_room])
        div_teams = [util.rpad([i for i, j in enumerate(pick_order) if j == div],
                               self.divisions*(3*math.ceil(max_room/3)*rooms - excess), None)
                     for div, (teams, rooms) in enumerate(self.divs)]

        team_order = [list(zip(idxs, teams)) for idxs, (teams, rooms) in zip(div_teams, self.divs)]
        self.teams = [team for idx, team in sorted(sum(list(team_order), []))]

        self.j_slots = [[], [], []]
        for teams, (_, rooms) in zip(div_teams, self.divs):
            rot_dir = 1 + bool(-len(teams) % 3 == 2)
            tmp = [sum([teams[(j + i*rot_dir) % 3 : len(teams) : 3] for i in range(3)],
                       [])[self.j_calib:] for j in range(3)]
            for i in range(3):
                self.j_slots[i] += [tmp[i][j::rooms] for j in range(rooms)]
        self.j_slots = [[util.rpad(room, max_room, None) for room in cat] for cat in self.j_slots]
        self.j_slots = list(zip(*[list(zip(*cat)) for cat in self.j_slots]))
        if self.j_calib:
            self.j_slots = [([0], [1], [2])] + self.j_slots
        self.assign_judge_times()

    def schedule_block(self):
        """Generates judging and table schedules using block scheduling."""
        raise NotImplementedError("Block scheduling is not implemented yet")

    def split_divisions(self):
        """Sets self.divs to a list of (teams, rooms for those teams) based on division."""
        if not self.divisions:
            self.divs = [(self.teams, self.j_sets)]
        room_max = max(12, math.ceil(self.num_teams / self.j_sets))

        self.divs = [(teams, math.ceil(len(teams) / room_max)) for teams in
                     [[team for team in self.teams if team.div == div] for div in
                      {team.div for team in self.teams}]]
        self.divs.sort(key=lambda div: -len(div[0]) % room_max)

        if sum([rooms for _, rooms in self.divs]) > self.j_sets:
            room_divs, impure_divs = [], []
            teams_left, rooms_left = self.num_teams, self.j_sets
            for teams, rooms in self.divs:
                if (-len(teams) // room_max + rooms_left) * room_max >= teams_left - len(teams):
                    room_divs.append((teams, rooms))
                    teams_left -= len(teams)
                    rooms_left -= rooms
                else:
                    impure_divs += [teams]

            room_max = math.ceil(teams_left / rooms_left)
            excess = rooms_left*room_max - teams_left

            #divisions that cnanot be isolated are simply run together and split into rooms
            #goals: no room with more than two divisions, as few split rooms as possible
            #split rooms are treated as separate divisions and therefore need to have fewer
            #teams per room than other divisions to avoid pointlessly idling judge rooms
            idx, spillover = 0, 0
            impure_teams = sum(impure_divs, [])
            for i, teams in enumerate(impure_divs):
                skips = int(i < excess % len(impure_divs)) + (excess // len(impure_divs))
                pure = ((len(teams) - spillover) // room_max) * room_max - max(0, skips - 1)
                room_divs.append((impure_teams[idx:idx + pure], math.ceil(pure / room_max)))

                idx += pure
                room_divs.append((impure_teams[idx:idx + room_max - bool(skips)], 1))
                idx += room_max - bool(skips)
                spillover += pure + room_max - bool(skips) - len(teams)
            self.divs = [(teams, rooms) for teams, rooms in room_divs if teams]

    def assign_judge_times(self):
        """Determines when each judging session will happen and assigns teams to those slots."""
        breaks = {0, len(self.j_slots)} | set(range(self.j_calib, len(self.j_slots), self.j_break[0]))
        breaks = sorted(list(breaks))
        times = [[self.j_start + bool(j and self.j_calib)*self.travel
                  + max(i - self.j_calib, 0)*self.j_break[1] + j*self.j_duration[0]
                  for j in range(breaks[i], breaks[i+1] + 1)] for i in range(len(breaks) - 1)]
        j_blockers = [self.j_lunch] + [(start - self.travel, duration + 2*self.travel) 
                      for start, duration in (self.opening, self.coach_meet)]
        for start, length in sorted(j_blockers):
            delay = max(timedelta(0), min(length, start + length - times[0][0]))
            delay -= self.j_break[1] if start >= times[0][-1] else timedelta(0)
            times = [[time + (start < cycle[-1])*delay for time in cycle] for cycle in times]
        times = [time for cycle in times for time in cycle]

        for breaktime in breaks[-2:0:-1]:
            self.j_slots.insert(breaktime, None)
        self.j_slots = list(zip(times, self.j_slots))
        for time, teams in filter(lambda x: x[1] is not None, self.j_slots):
            for cat, cat_teams in enumerate(teams):
                for room, team in filter(lambda x: x[1] is not None, enumerate(cat_teams)):
                    self.teams[team].add_event(time, self.j_duration[1], cat + 2, room)

    def schedule_matches(self, time, team, rounds, run_rate):
        """Determines when table matches will occur and assigns teams to matches."""
        def avail(start, rnd):
            return [(t, self._team(start + t).available(time, self.t_duration[rnd], self.travel))
                    for t in range(self.num_teams)]
        #it isn't always possible to run tables at full speed; we prefer to consistently idle a few
        #tables over idling all tables for a match
        start_team = team
        if run_rate is None:
            match_sizes = [2*self.t_pairs]
        else:
            run_rate = 2*min(math.ceil(run_rate/2), self.t_pairs)
            match_sizes = [max(2, run_rate - 2), run_rate]

        while team < len(rounds)*self.num_teams + start_team:
            rnd = rounds[(team - start_team) // self.num_teams]
            max_teams = next((t for t, free in avail(team, rnd) if not free),
                             len(rounds)*self.num_teams)
            max_matches = (self._team(team).next_event(time)[0] - time - self.travel)\
                    // self.t_duration[rnd]

            if team + max_teams >= len(rounds)*self.num_teams + start_team:
                max_teams = len(rounds)*self.num_teams + start_team - team
                max_matches = min(max_matches, math.ceil(max_teams / match_sizes[-1]))
                min_matches = math.ceil(max_teams / match_sizes[-1])
            else:
                max_matches = min(math.ceil((len(rounds)*self.num_teams + start_team - team)
                                            / match_sizes[-1]), max_matches)
                min_matches = math.ceil((self._team(team + max_teams).next_available(time, self.travel)
                                         - time) / self.t_duration[rnd])
                max_teams -= max_teams % 2

            if max_matches == 0:
                time += self.t_duration[rnd]

            next_matches = util.alt_sum_to(match_sizes, max_teams, min_matches, force_take_all=
                                       team + max_teams - start_team >= len(rounds)*self.num_teams)
            while sum(next_matches[:-1]) + (team - start_team) % self.num_teams > self.num_teams:
                next_matches.pop()

            #schedule as many teams as we can in the currently available team and time block
            for match_size in next_matches:
                timeslot = [t % self.num_teams for t in range(team, team + match_size)]
                self.t_slots += [(time, rnd, util.rpad(timeslot, match_sizes[-1], None))]
                team += match_size
                time += self.t_duration[rnd]

        return time

    def assign_tables(self, assignment_passes=2):
        """Reorders the teams in self.t_slots to minimize table repetition for teams."""
        prev_tables = [[0 for i in range(2*self.t_pairs)] for j in range(self.num_teams)]
        def cost(order):
            val = sum(prev_tables[team][table]**1.1 for table, team in enumerate(order)
                      if team is not None)
            unpaired = sum(1 for i in range(0, len(order) - 1, 2) if
                           (order[i] == None) != (order[i + 1] == None))
            val += unpaired * (max((max(row) for row in prev_tables)) + 1)
            return val

        #the current approach only changes one match at a time; multiple passes fix bad early calls
        for assign_pass in range(assignment_passes):
            rotation = 0
            for(time, rnd, teams) in filter(None, self.t_slots):
                if assign_pass:
                    for table, team in filter(lambda x: x[1] is not None, enumerate(teams)):
                        prev_tables[team][table] -= 1
                else:
                    rotation = (rotation + sum(1 for team in teams if team is None)) % len(teams)
                    rotation -= rotation % 2
                teams[:] = scheduler.min_cost.min_cost(teams[rotation:] + teams[:rotation], cost)
                for table, team in filter(lambda x: x[1] is not None, enumerate(teams)):
                    prev_tables[team][table] += 1
                    if assign_pass + 1 == assignment_passes:
                        teams[:] = util.rpad(teams, 2*self.t_pairs, None)
                        team_rnd = sum(1 for event in self._team(team).events if event[2] > 4)
                        self._team(team).add_event(time, self.t_duration[rnd], 5 + team_rnd, table)

    def _team(self, team_num):
        """Returns the team at the specified internal index; wraps modularly."""
        return self.teams[team_num % self.num_teams]
