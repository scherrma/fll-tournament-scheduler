#!/usr/bin/python3
import pandas
import math
from datetime import datetime, timedelta
from sched.team import Team
import sched.util #chunks
from sched.util import clean_input, strtotime, positive

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
                try_again = clean_input("Would you like to try another file (y/n): ",
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
        self.j_cycle_len = 3

        self.travel_time = timedelta(minutes=12.5)

        self.t_pairs = round(3 / 4 * self.j_sets)
        self.t_duration_early = timedelta(minutes=10)
        self.t_duration_lunch = timedelta(minutes=45)
        self.t_duration_late = timedelta(minutes=8)

    def schedule_interleaved(self):
        #judge slot scheduling
        self.j_slots = [sum([list(range((i + j) % 3, self.num_teams, 3))
                            for i in range(3)], []) for j in range(3)]
        self.j_slots = list(zip(*[sched.util.chunks(l[self.j_calib:], self.j_sets) 
                                      for l in self.j_slots]))
        if self.j_calib:
            self.j_slots = [([0], [1], [2])] + self.j_slots

        for timeslot in range(len(self.j_slots)):
            time = self.j_start + timeslot*self.j_duration\
                 + (timeslot + 2*self.j_calib) // self.j_cycle_len * self.j_break

            for cat in range(3):
                for i in range(len(self.j_slots[timeslot][cat])):
                    self.teams[self.j_slots[timeslot][cat][i]]\
                            .add_event(time, self.j_duration_team, self.j_names[cat], i)
        
        #table scheduling
        t_run_rate = min(self.t_pairs, 3/2*self.j_sets * self.t_duration_early
                        /(self.j_duration + self.j_break/self.j_cycle_len))
        t_match_sizes = [2*math.floor(t_run_rate), 2*math.ceil(t_run_rate)]

       # for team in self.teams:
       #     team.add_event(team.events[0][0] + team.events[0][1] + self.travel_time, 
       #             self.t_duration_early, self.t_names[0], 2)

       # team_idx = 0
       # start_time = self.teams[3 * self.j_calib]

       # while team_idx < 2*self.num_teams:
       #     rd = team_idx // self.num_teams
       #     team = team_idx % self.num_teams
             
               
    def export(self):
        for team in self.teams:
            print(team, "\tclosest:",team.closest())
            print(''.join(['\t' + time.strftime('%r') + ' - ' + name + ' ' +
                str(loc) + '\n' for (time, d, name, loc) in team.events]))
            
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
                sheet = clean_input("Which sheet has the team list: "
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
