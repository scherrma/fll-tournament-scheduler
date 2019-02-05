#!/usr/bin/python3
import pandas
import math
from datetime import datetime, timedelta
from sched.team import Team
import sched.util as util

class Tournament:
    j_names = ['Project', 'Robot Design', 'Core Values'] #judging category names
    t_names = ['Practice', 'Round 1', 'Round 2', 'Round 3'] #table round category names
    
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

        self.travel_time = timedelta(minutes=12.5)

        self.t_pairs = round(3 / 4 * self.j_sets)
        self.t_duration_early = timedelta(minutes=10)
        self.t_duration_lunch = timedelta(minutes=45)
        self.t_duration_late = timedelta(minutes=8)

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
        t_team_rate = min(2*self.t_pairs, 3*self.j_sets*self.t_duration_early 
                            / (self.j_duration + self.j_break/self.j_consec))
        t_team_options = [2*math.ceil(t_team_rate/2 - 1), 2*math.ceil(t_team_rate/2)]

        avail_at = lambda team, time: self.teams[team % self.num_teams]\
                .available(time, self.t_duration_early, self.travel_time)

        time_start = self.teams[3*self.j_calib].events[0][0]\
                   + self.teams[3*self.j_calib].events[0][1] + self.travel_time
        team_idx = util.find_first(range(3*self.j_calib + 1), lambda x: avail_at(x, time_start))

        remainder = 2*(self.num_teams % self.t_pairs)
        if all([avail_at((team_idx - i) % self.num_teams, time_start) for i in range(remainder)]):
            team_idx = (team_idx - remainder) % self.num_teams
            time_start -= self.t_duration_early

        time = self.schedule_tables(time_start, team_idx, self.t_duration_early, t_team_options)[0]
        self.schedule_tables(time + self.t_duration_lunch, team_idx, self.t_duration_late, t_team_options)

    def schedule_tables(self, start_time, start_team, duration, match_sizes):
        time, team = start_time, start_team
        while team < 2*self.num_teams + start_team:
            try:
                max_teams = next((t for t in range(self.num_teams) if not 
                    self.teams[(t + team) % self.num_teams].available(time, duration, self.travel_time)))
            except:
                max_teams = 2*self.num_teams
            max_matches = (self.teams[team % self.num_teams].next_event(time)[0] - time - self.travel_time) // duration
            if team + max_teams >= 2*self.num_teams + start_team:
                max_teams = 2*self.num_teams + start_team - team
                max_matches = math.ceil(max_teams / match_sizes[-1]) 

            for match_size in util.sum_to(match_sizes, max_teams, max_matches):
                for t in range(team, team + match_size):
                    self.teams[t % self.num_teams].add_event(time, duration, -1, -1)
                team += match_size
                time += duration
        return (time, team)
               
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
        
if __name__ == "__main__":
    Tournament().schedule()
