#!/usr/bin/env python3
import pandas as pd
import math
from datetime import datetime, timedelta
from scheduler.team import Team
import scheduler.util as util

import xlrd

import tkinter
from tkinter import filedialog
import sys
import os

class Tournament:
    def __init__(self):
        if len(sys.argv) == 1:
            root = tkinter.Tk()
            root.withdraw()
            self.fpath = filedialog.askopenfilename(
                    filetypes=[("Excel files (*.xls, *.xlsm, *.xlsx)", "*.xls;*.xlsm;*.xlsx")])
            root.destroy()
        elif len(sys.argv) == 2:
            self.fpath = sys.argv[1]
        else:
            print("{} accepts 0 or 1 arguments; {} received".format(sys.argv[0], len(sys.argv) - 1))

    def schedule(self):
        try:
            self.read_data(self.fpath)

            if self.scheduling_method == "Interlaced":
                self.schedule_interlaced()
            elif self.scheduled_method == "Block":
                self.schedule_block()
            else:
                raise ValueError(self.scheduling_method + " scheduling is not supported")

            self.export()

        except Exception as e:
            raise e
            print(e)
            if sys.platform == "win32":
                os.system("pause")
            raise SystemExit
    
    def schedule_interlaced(self):
        #judge slot scheduling
        rot_dir = 1 if (3*self.j_calib - self.num_teams) % 3 is 1 else -1
        self.j_slots = [sum([list(range((j + i*rot_dir) % 3, self.num_teams, 3))
                            for i in range(3)], []) for j in range(3)]
        self.j_slots = list(zip(*[util.chunks(l[self.j_calib:], self.j_sets) 
                                      for l in self.j_slots]))
        if self.j_calib:
            self.j_slots = [([0], [1], [2])] + self.j_slots

        time = self.j_start
        for timeslot in range(len(self.j_slots)):
            if (timeslot - self.j_calib) % self.j_consec == 0 and 0 < timeslot < len(self.j_slots) - 1:
                time += self.j_break
            for cat in range(3):
                for i in range(len(self.j_slots[timeslot][cat])):
                    self.teams[self.j_slots[timeslot][cat][i]]\
                            .add_event(time, self.j_duration_team, cat, i)
            time += self.j_duration

        #table scheduling
        time_start = sum(self._team(3*self.j_calib).events[0][:2], self.travel)
        team_idx = next((t for t in range(3*self.j_calib + 1) if
            self._team(t).available(time_start, self.t_duration[0], self.travel)))

        round_split = [(0, 1), (2, 3)]
        run_rates = [3*self.j_sets*self.t_duration[0]/(self.j_duration + self.j_break/self.j_consec), None]
        break_times = (None, self.t_lunch_duration)

        min_early_run_rate = max(2, 2*math.ceil(run_rates[0]/2))
        if self.num_teams % min_early_run_rate:
            early_avail = [self._team(team_idx - i).available(time_start - self.t_duration[0],
                self.t_duration[0], self.travel) for i in range(min_early_run_rate)]
            if all(early_avail):
                team_idx -= len(early_avail)
                time_start -= self.t_duration[0]

        self.schedule_tables(time_start, team_idx, zip(round_split, run_rates, break_times))

    def schedule_block(self):
        raise NotImplementedError

    def schedule_tables(self, time, team, round_info):
        start_team = team
        for (rounds, run_rate, break_time) in round_info:
            if run_rate is None or run_rate > 2*self.t_pairs:
                run_rate = 2*self.t_pairs
            run_rate = 2*math.ceil(run_rate/2)
            match_sizes = [max(2, run_rate - 2), run_rate]

            time += (break_time if break_time else timedelta(0))
            team = next((t for t in range(team, team + self.num_teams) if
                        self._team(t).available(time, self.t_duration[rounds[0]], self.travel)))
            start_team = team

            while team < 2*self.num_teams + start_team:
                rd = rounds[(team - start_team) // self.num_teams]
                max_teams = next((t for t in range(self.num_teams) if not self._team(t + team)
                    .available(time, self.t_duration[rd], self.travel)), 2*self.num_teams)
                max_matches = (self._team(team).next_event(time)[0] - time - self.travel)\
                        // self.t_duration[(team - start_team) // self.num_teams]

                if team + max_teams >= 2*self.num_teams + start_team:
                    max_teams = 2*self.num_teams + start_team - team
                    max_matches = math.ceil(max_teams / match_sizes[-1])

                for match_size in util.sum_to(match_sizes, max_teams, max_matches):
                    for t in range(team, team + match_size):
                        self._team(t).add_event(time, self.t_duration[rd], 3 
                                + rounds[(t - start_team) // self.num_teams], t % self.t_pairs)
                    team += match_size
                    time += self.t_duration[rd]
        return time

               
    def export(self):
        longest_name = max([len(str(team)) for team in self.teams])
        print(self.rooms)
        for team in self.teams:
            print("{:<{}}   (closest: {})".format(str(team), longest_name, team.closest()))
            print(''.join(['\t{} - {} ({})\n'.format(time.strftime('%r'), self.event_names[cat], 
                            self.rooms[min(cat, 3)][loc]) for (time, dur, cat, loc) in team.events]))
            
    def _team(self, team_num):
        return self.teams[team_num % self.num_teams]

    def read_data(self, fpath):
        dfs = pd.read_excel(fpath, sheet_name=["Team Information", "Input Form"])

        team_sheet = dfs["Team Information"]
        if all([x in team_sheet.columns for x in ("Team Number", "Team")]):
            self.teams = [Team(*x) for x in team_sheet.loc[:, ("Team Number", "Team")].values]
        else:
            raise KeyError("Could not find columns 'Team Number' and 'Team' in sheet 'Team Information'")

        param_sheet = dfs["Input Form"]
        if any([x not in param_sheet.columns for x in ("key", "answer")]):
            raise KeyError("Could not find columns 'key' and 'answer' in sheet 'Input Form'")
        param_sheet = param_sheet.set_index("key")
        param = dict(param_sheet.loc[:, "answer"].items())

        self.num_teams = len(self.teams)
        try:
            self.tournament_name = param["tournament_name"]
            self.scheduling_method = param["scheduling_method"]
            self.travel = timedelta(minutes=param["travel_time"])
            self.event_names = ['Project', 'Robot Design', 'Core Values']\
                             + param_sheet.loc["t_round_names"].dropna().values.tolist()[1:]
            self.rooms = [param_sheet.loc[key].dropna().values.tolist()[1:]
                         for key in ("j_project_rooms", "j_robot_rooms", "j_values_rooms")]
            self.rooms += [sum([[tbl + ' A', tbl + ' B'] for tbl in param_sheet.loc["t_pair_names"]
                                .dropna().values.tolist()[1:]], [])]
            
            self.j_start = datetime(1, 1, 1, param["j_start"].hour, param["j_start"].minute)
            self.j_sets = param["j_sets"]
            self.j_calib = (param["j_calib"] == "Yes")
            self.j_duration = timedelta(minutes=param["j_duration"])
            self.j_duration_team = timedelta(minutes=10)
            self.j_consec = param["j_consec"]
            self.j_break = timedelta(minutes=param["j_break"])
            self.j_lunch = datetime(1, 1, 1, param["j_lunch"].hour, param["j_lunch"].minute)
            self.j_lunch_duration = timedelta(minutes=param["j_lunch_duration"])
            
            self.t_pairs = param["t_pairs"]
            self.t_rounds = len(self.event_names[-1])
            self.t_duration = [timedelta(minutes=x) for x in 
                param_sheet.loc["t_durations"].dropna().values.tolist()[1:]]
            self.t_lunch = datetime(1, 1, 1, param["t_lunch"].hour, param["t_lunch"].minute)
            self.t_lunch_duration = timedelta(minutes=param["t_lunch_duration"])
            
        except KeyError as e:
            raise KeyError(str(e) + " not found in 'key' column in sheet 'Input Form'")
        
if __name__ == "__main__":
    Tournament().schedule()
