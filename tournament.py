#!/usr/bin/python3
import pandas
import math
from datetime import datetime, timedelta
from sched.team import Team
import sched.util #chunks
from sched.util import clean_input, strtotime, positive

class Tournament:
    j_names = ['Project', 'Robot Design', 'Core Values']
    t_names = ['Practice', 'Round']
    
    def schedule(self):
        self.setup()
        self.assign_judging()
        self.assign_tables()

        for team in self.teams:
            activities = [(time, Tournament.j_names[cat] + ' ' + str(loc + 1))
                    if cat >= 0 else (time, 'Table ' + Tournament.t_names[bool(-cat-1)])
                          for (time, cat, loc) in team.events]
            activities.sort()
            print('\n' + str(team) + 
                    ''.join(['\n\t' + loc + ' at ' + time.strftime('%I:%M%p') for (time, loc) in activities]))

    def assign_judging(self):
        self.j_slots = [sum([list(range((i + j) % 3, len(self.teams), 3))
                            for i in range(3)], []) for j in range(3)]
        self.j_slots = list(zip(*[sched.util.chunks(l[self.calib:], self.j_sets) 
                                      for l in self.j_slots]))
        if self.calib:
            self.j_slots = [([0], [1], [2])] + self.j_slots

        for timeslot in range(len(self.j_slots)):
            time = self.j_start + timeslot*self.j_duration\
                 + (timeslot + 2*self.calib) // self.j_cycle_len * self.j_break

            for cat in range(3):
                for i in range(len(self.j_slots[timeslot][cat])):
                    self.teams[self.j_slots[timeslot][cat][i]].events.append((time, cat, i))

    def assign_tables(self):
        t_start = sorted(self.teams[0].events)[1][0] - self.cross_time - self.t_duration_early
        t_early = [2 * int(3 * self.j_sets * i / 4) for i in
                range(math.ceil(4 * len(self.teams) / (3 * self.j_sets)))] + [2*len(self.teams)]

        for i in range(len(t_early) - 1):
            print((t_start + i*self.t_duration_early).strftime('%r'), "teams:", 
                    list(map(lambda x: x % len(self.teams), range(t_early[i], t_early[i+1]))))
            for team in range(t_early[i], t_early[i + 1]):
                self.teams[team % len(self.teams)].events.append((t_start + i*self.t_duration_early,
                        -(t_early[i] + team) // len(self.teams) - 1, team))

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
        if good_data:
            print(len(self.teams), "teams read")
            self.set_params()
                      
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
            df = xl.parse(sheet_name = sheet)
        
        try:
            self.teams = [Team(*x) for x in df.filter(items=["Team Number", "Team"]).values]
        except TypeError:
            raise IOError("Could not find find columns 'Team Number' and 'Team'.")

    def set_params(self):
        self.j_sets = math.ceil(len(self.teams) / 12)
        self.calib = (self.j_sets > 2) or (len(self.teams) % self.j_sets == 1)
        self.j_start = datetime(1, 1, 1, 9)
        self.j_duration = timedelta(minutes=17.5)
        self.j_break = timedelta(minutes=7.5)
        self.j_cycle_len = 3

        self.cross_time = timedelta(minutes=12.5)

        self.t_pairs = round(3 / 4 * self.j_sets)
        self.t_duration_early = timedelta(minutes=10)
        self.t_duration_lunch = timedelta(minutes=45)
        self.t_duration_late = timedelta(minutes=8)

if __name__ == "__main__":
    Tournament().schedule()
