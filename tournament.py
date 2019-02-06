#!/usr/bin/env python3
import pandas
import math
from datetime import datetime, timedelta
from sched.team import Team
import sched.util as util

class Tournament:
    def schedule(self):
        self.setup()
        self.set_params()
        self.schedule_interleaved()
        self.export()
    
    def setup(self):
        good_data = False
        try_again = 'y'
        while try_again == 'y' and not good_data:
            try:
                self.ingest_sheet('/mnt/c/Users/Marty/Desktop/fll-tournament-scheduler/testdata/example_team_list.xlsx')
                good_data = True
            except IOError as error:
                print(error)
                try_again = util.clean_input("Would you like to try another file (y/n): ",
                                        lambda x: x in ('y', 'n'), lambda x: x.lower())
        print(len(self.teams), "teams read")
   
    def set_params(self):
        self.num_teams = len(self.teams)
        self.j_sets = math.ceil(self.num_teams / 12)
        self.j_calib = (self.j_sets > 2) or (self.num_teams % self.j_sets == 1)
        self.j_start = datetime(1, 1, 1, 9)
        self.j_duration = timedelta(minutes=17.5)
        self.j_duration_team = timedelta(minutes=10)
        self.j_break = timedelta(minutes=7.5)
        self.j_consec = 3
        self.j_names = ['Project', 'Robot Design', 'Core Values']

        self.travel = timedelta(minutes=12.5)

        self.t_pairs = round(3 / 4 * self.j_sets)
        self.t_duration = 2*[timedelta(minutes=10)] + 2*[timedelta(minutes=8)]
        self.t_lunch = timedelta(minutes=45)
        self.t_names = ['Practice', 'Round 1', 'Round 2', 'Round 3']

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
                            .add_event(time, self.j_duration_team, self.j_names[cat], i)
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
        self.schedule_tables(time + self.t_lunch, team_idx, (2, 3))

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
                    self._team(t).add_event(time, self.t_duration[rd], -1, -1)
                team += match_size
                time += self.t_duration[rd]
        return time
               
    def export(self):
        longest_name = max([len(str(team)) for team in self.teams])
        for team in self.teams:
            print("{:<{}}   (closest: {})".format(str(team), longest_name, team.closest()))
            print(''.join(['\t{} - {} {} ({})\n'.format(time.strftime('%r'), name,
                loc + 1, duration) for (time, duration, name, loc) in team.events]))
            
    def ingest_sheet(self, filepath):
        filetype = filepath.split('.')[-1]
        if filetype not in ('csv', 'xls', 'xlsx'):
            raise IOError("Invalid file type. The file must be either csv, xls, or xlsx.")
            
        if filetype == 'csv':
            try:
                df = pandas.read_csv(filepath)
            except pandas.errors.EmptyDataError:
                raise IOError("The selected file is empty.")
        elif filetype in ['xls', 'xlsx']:
            sheet = 0
            xl = pandas.ExcelFile(filepath)
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
        
if __name__ == "__main__":
    Tournament().schedule()
