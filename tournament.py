#!/usr/bin/python3
import pandas
import math
from datetime import datetime, timedelta
from sched.team import Team
from sched.helpers import clean_input, strtotime, positive

class Tournament:
    def schedule(self):
        self.setup()

    def setup(self):
        good_data = False
        try_again = 'y'
        while try_again == 'y' and not good_data:
            try:
                self.ingest_sheet('/mnt/c/Users/Marty/Desktop/fll scheduler/example fll team list.xlsx')
                good_data = True
            except IOError as error:
                print(error)
                try_again = clean_input("Would you like to try another file (y/n): ",
                                        lambda x: x in ('y', 'n'), lambda x: x.lower())
        if good_data:
            print(len(self.teams), "teams read")
            self.get_params()
                      
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

    def get_params(self):
        self.judge_rooms = math.ceil(len(self.teams) / 12)
        self.calib = (self.judge_rooms > 2) or (len(self.teams) % self.judge_rooms == 1)
        self.judge_length = math.ceil((len(self.teams) - self.calib) / self.judge_rooms)
        self.judge_start = datetime(1, 1, 1, 9)
        self.judge_time = max(timedelta(minutes=17.5), min(timedelta(minutes=20),
                              (datetime(1, 1, 1, 13) - self.judge_start) / self.judge_length))

        print("{} judge rooms, seeing {} teams each\nStarting at {}, "
                "seeing each team for {} and ending at {}".format(
                    self.judge_rooms, self.judge_length, self.judge_start.strftime('%I:%M%p'),
                    self.judge_time, (self.judge_start + (self.judge_length + 1) * self.judge_time)
                    .strftime('%I:%M%p')))

       # use_defaults = clean_input("Would you like to use the default settings (y/n): ",
       #                            lambda x: x in ('y', 'n'), lambda x: x.lower()) == 'y'
       # self.judge_rooms = math.ceil(len(self.teams) / 12)
       # if not use_defaults:
       #     self.judge_rooms = clean_input("How many judging rooms are there "
       #         "for each category (default {}): ".format(self.judge_rooms),
       #         positive, int)

       # self.judge_start = datetime(1, 1, 1, 9) #9am
       # if not use_defaults:
       #     self.judge_start = clean_input("When will judging start (hh:mm) (default {}): "
       #                        .format(self.judge_start.strftime('%I:%M%p')), parse=strtotime)

       # self.judge_time = max(timedelta(minutes=17.5), min(timedelta(minutes=20),
       #                       (datetime(1, 1, 1, 13) - self.judge_start)))
       # print("self.judge_time:", self.judge_time)
        #if not use_defaults:
        #    self.judge_time = timedelta(minutes=clean_input("How many "
        #        "minutes will the judges have for each team (default {}): "
        #        .format(self.judge_time), positive, float))


if __name__ == "__main__":
    Tournament().schedule()
