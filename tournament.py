#!/usr/bin/env python3
import pandas as pd
import math
from datetime import datetime, timedelta
from scheduler.team import Team
import scheduler.util as util

class Tournament:
    def __init__(self):
        self.fpath = "input_sheet.xlsx"

    def schedule(self):
        self.read_data(self.fpath)
        self.schedule_interleaved()
        self.export()
    
    def schedule_interleaved(self):
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
            if (timeslot - self.j_calib) % self.j_consec == 0 and timeslot + 1 < len(self.j_slots):
                time += self.j_break
            for cat in range(3):
                for i in range(len(self.j_slots[timeslot][cat])):
                    self.teams[self.j_slots[timeslot][cat][i]]\
                            .add_event(time, self.j_duration_team, cat, i)
            time += self.j_duration

        #table scheduling
        time_start = sum(self._team(3*self.j_calib).events[0][:2], self.travel)
        team_idx = next((t for t in range(3*self.j_calib) if
            self._team(t).available(time_start, self.t_duration[0], self.travel)))

        early_avail = [self._team(team_idx - i).available(time_start - self.t_duration[0],
                self.t_duration[0], self.travel) for i in range(2*(self.num_teams % self.t_pairs))]
        if all(early_avail):
            team_idx -= len(early_avail)
            time_start -= self.t_duration[0]

        early_run_rate = 3*self.j_sets*self.t_duration[0]/(self.j_duration + self.j_break/self.j_consec)
        time = self.schedule_tables(time_start, team_idx, (0, 1), early_run_rate)
        team_idx = next((t for t in range(team_idx, team_idx + self.num_teams) if
                        self._team(t).available(time, self.t_duration[2], self.travel)))
        self.schedule_tables(time + self.t_lunch_duration, team_idx, (2, 3))

    def schedule_tables(self, start_time, start_team, rounds, run_rate=None):
        if run_rate is None or run_rate > 2*self.t_pairs:
            run_rate = 2*self.t_pairs
        run_rate = 2*math.ceil(run_rate/2)
        match_sizes = [run_rate - 2, run_rate]

        time, team = start_time, start_team
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
                    self._team(t).add_event(time, self.t_duration[rd], 3, 0)
                team += match_size
                time += self.t_duration[rd]
        return time
               
    def export(self):
        longest_name = max([len(str(team)) for team in self.teams])
        for team in self.teams:
            print("{:<{}}   (closest: {})".format(str(team), longest_name, team.closest()))
            print(''.join(['\t{} - {} ({})\n'.format(time.strftime('%r'), self.event_names[cat],
                self.rooms[cat][loc]) for (time, dur, cat, loc) in team.events]))
            
    def ingest_sheet(self, filepath):
        sheet = 0
        xl = pd.ExcelFile(filepath)
        if len(xl.sheet_names) > 1:
            sheets = list(map(lambda x: x.lower(), xl.sheet_names))
            sheet = util.clean_input("Which sheet has the team list: "
                        + ', '.join(xl.sheet_names) + '\n',
                        parse = lambda x: sheets.index(x.lower()))
        df = xl.parse(sheet_name=sheet)
        
        try:
            self.teams = [Team(num, name) for 
                    (num, name) in df.filter(items=["Team Number", "Team"]).values]
        except TypeError:
            raise IOError("Could not find find columns 'Team Number' and 'Team'.")

    def _team(self, team_num):
        return self.teams[team_num % self.num_teams]

    def read_data(self, fpath):
        team_sheet = pd.read_excel(fpath, sheet_name="Team Information")
        self.teams = [Team(*x) for x in team_sheet.loc[:, ("Team Number", "Team")].values]

        param_sheet = pd.read_excel(fpath, sheet_name="Input Form").set_index("key")
        param = dict(param_sheet.loc[:, "answer"].items())

        self.num_teams = len(self.teams)
        self.tournament_name = param["tournament_name"]
        self.scheduling_method = param["scheduling_method"]
        self.travel = timedelta(minutes=param["travel_time"])

        self.j_start = datetime(1, 1, 1, param["j_start"].hour, param["j_start"].minute)
        self.j_sets = param["j_sets"]
        self.j_calib = (param["j_calib"] == "Yes")
        self.j_duration = timedelta(minutes=param["j_duration"])
        self.j_consec = param["j_consec"]
        self.j_break = timedelta(minutes=param["j_break"])
        self.j_lunch = datetime(1, 1, 1, param["j_lunch"].hour, param["j_lunch"].minute)
        self.j_lunch_duration = timedelta(minutes=param["j_lunch_duration"])

        self.rooms = [param_sheet.loc[key].dropna().values.tolist()[1:]
                        for key in ("j_project_rooms", "j_robot_rooms", "j_values_rooms")]
        
        #these are hacks and should be fixed or done elsewhere
        self.rooms += [["Competition Floor"]]
        self.event_names = ['Project', 'Robot Design', 'Core Values', 'Robot Game']
        self.j_duration_team = timedelta(minutes=10)

        #back to not-hacks
        self.t_pairs = param["t_pairs"]
        self.t_lunch = datetime(1, 1, 1, param["t_lunch"].hour, param["t_lunch"].minute)
        self.t_lunch_duration = timedelta(minutes=param["t_lunch_duration"])

        self.t_round_names = param_sheet.loc["t_round_names"].dropna().values.tolist()[1:]
        self.t_rounds = len(self.t_round_names)
        self.t_duration = [timedelta(minutes=x) for x in 
                param_sheet.loc["t_durations"].dropna().values.tolist()[1:]]
        
if __name__ == "__main__":
    Tournament().schedule()
