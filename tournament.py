#!/usr/bin/env python3
import pandas as pd
import math
import numpy as np
from datetime import datetime, timedelta
from scheduler.team import Team
import scheduler.util as util

import openpyxl
import openpyxl.styles as styles

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
            self.assign_tables()
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

        breaks = []
        time = self.j_start
        for timeslot in range(len(self.j_slots)):
            if (timeslot - self.j_calib) % self.j_consec == 0 and 0 < timeslot < len(self.j_slots) - 2:
                time += self.j_break
                breaks.append(timeslot)
            for cat in range(3):
                for i in range(len(self.j_slots[timeslot][cat])):
                    self.teams[self.j_slots[timeslot][cat][i]]\
                            .add_event(time, self.j_duration_team, cat, i)
            self.j_slots[timeslot] = (time, list(self.j_slots[timeslot]))
            time += self.j_duration
        for (time, areas) in self.j_slots[self.j_calib:]:
            for area in areas:
                area += (self.j_sets - len(area)) * [None]
        for breaktime in breaks[::-1]:
            self.j_slots.insert(breaktime, None)
        

        #table scheduling
        time_start = sum(self._team(3*self.j_calib).events[0][:2], self.travel)
        team_idx = next((t for t in range(3*self.j_calib + 1) if
            self._team(t).available(time_start, self.t_duration[0], self.travel)))

        round_split = [(0, 1), (2, 3)]
        run_rates = [3*self.j_sets*self.t_duration[0]/(self.j_duration + self.j_break/self.j_consec), None]
        break_times = (self.t_lunch_duration, None)

        min_early_run_rate = max(2, 2*math.ceil(run_rates[0]/2))
        if self.num_teams % min_early_run_rate:
            early_avail = [self._team(team_idx - i).available(time_start - self.t_duration[0],
                self.t_duration[0], self.travel) for i in range(min_early_run_rate)]
            if all(early_avail):
                team_idx -= len(early_avail)
                time_start -= self.t_duration[0]

        self.schedule_matches(time_start, team_idx, zip(round_split, run_rates, break_times))

    def schedule_block(self):
        raise NotImplementedError

    def schedule_matches(self, time, team, round_info):
        self.t_slots = []
        t_idle = 0
        start_team = team
        for (rounds, run_rate, break_time) in round_info:
            if run_rate is None or run_rate > 2*self.t_pairs:
                run_rate = 2*self.t_pairs
            run_rate = 2*math.ceil(run_rate/2)
            match_sizes = [max(2, run_rate - 2), run_rate]

            team = next((t for t in range(team, team + self.num_teams) if
                        self._team(t).available(time, self.t_duration[rounds[0]], self.travel)))
            start_team = team

            while team < 2*self.num_teams + start_team:
                rd = rounds[(team - start_team) // self.num_teams]
                max_teams = next((t for t in range(self.num_teams) if not self._team(t + team)
                    .available(time, self.t_duration[rd], self.travel)), 2*self.num_teams)

                if team + max_teams >= 2*self.num_teams + start_team:
                    max_teams = 2*self.num_teams + start_team - team
                    max_matches = min(max_matches, math.ceil(max_teams / match_sizes[-1]))
                elif self._team(team).next_event(time)[0] == datetime.max:
                    max_matches = max_teams // match_sizes[-1]
                else:
                    max_matches = (self._team(team).next_event(time)[0] - time - self.travel)\
                        // self.t_duration[rd]

                for match_size in util.sum_to(match_sizes, max_teams, max_matches):
                    timeslot = [i % self.num_teams for i in range(team, team + match_size)]
                    if match_size < match_sizes[-1]:
                        timeslot[t_idle:t_idle] = (match_sizes[-1] - match_size)*[None]
                        t_idle = (t_idle + 2) % match_sizes[-1]
                    timeslot += (2*self.t_pairs - match_sizes[-1])*[None]
                    self.t_slots += [(time, rd, timeslot)]
                    team += match_size
                    time += self.t_duration[rd]
            if break_time is not None:
                time += break_time
                self.t_slots += [None]
    
    def assign_tables(self):
        for (time, rd, teams) in [x for x in self.t_slots if x is not None]:
            for i in [j for j in range(len(teams)) if teams[j] is not None]:
                self._team(teams[i]).add_event(time, self.t_duration[rd], 3, i)
        for t in self.teams:
            rd = 0
            for event in [e for e in t.events if e[2] == 3]:
                event[2] += rd
                rd += 1
               
    def export(self):
        time_fmt_str = ('%#I:%M %p' if sys.platform == "win32" else '%-I:%M %p')
        wb = openpyxl.load_workbook(self.fpath)
        del wb["Input Form"]
        self._export_judge_views(wb, time_fmt_str)
        self._export_table_views(wb, time_fmt_str)
        self._export_team_views(wb, time_fmt_str)
        wb._sheets = wb._sheets[1:] + wb._sheets[:1]
               
        saved = False
        count = 0
        while not saved:
            outfpath = [os.path.join(os.path.dirname(self.fpath), 
                self.tournament_name.lower().replace(' ', '_') + '_schedule'), '.xlsx']
            try:
                wb.save((' ({})'.format(count) if count else '').join(outfpath))
                saved = True
            except PermissionError:
                count += 1
        print("file saved as", (' ({})'.format(count) if count else '').join(outfpath))

    def _export_judge_views(self, wb, time_fmt_str):
        thin = styles.Side(border_style='thin', color='000000')
        thick = styles.Side(border_style='thick', color='000000')

        ws_overall = wb.create_sheet("Judging Rooms")
        ws_overall.append(sum([[cat] + (2*self.j_sets - 1)*[''] for cat in self.event_names[:3]], ['']))
        ws_overall.append(sum([['', room] for room in sum(self.rooms[:3], [])], []))

        cat_sheets = [wb.create_sheet(cat) for cat in self.event_names[:3]]
        for i in range(3):
            cat_sheets[i].append([''] + [self.event_names[i]])
            cat_sheets[i].append(sum([['', room] for room in self.rooms[i]], []))

        for slot in self.j_slots:
            if slot is None:
                for ws in [ws_overall] + cat_sheets:
                    ws.append([''])
            else:
                time, teams = slot[0], [[None if t is None else self.teams[t] for t in cat] 
                                        for cat in slot[1]]
                if len(teams[0]) == 1 and self.j_calib:
                    line = sum([[teams[i][0].num, teams[i][0].name, "all {} judges in {}".format(
                        self.event_names[i], self.rooms[i][0])] + (2*self.j_sets - 3)*[''] for i 
                        in range(3)], [time.strftime(time_fmt_str)])
                else:
                    line = [time.strftime(time_fmt_str)] + sum([['', 'None'] if team is None 
                        else [team.num, team.name] for cat in teams for team in cat], [])
                ws_overall.append(line)
                for i in range(3):
                        cat_sheets[i].append([line[0]] + line[2*i*self.j_sets + 1:
                                                              2*(i + 1)*self.j_sets + 1]) 

        for ws in [ws_overall] + cat_sheets:
            for row in ws.rows:
                for cell in row:
                    cell.alignment = styles.Alignment(horizontal='center')
                    if cell.column % (2*self.j_sets) == 2:
                        cell.border = styles.Border(left=thick)
                    elif cell.column % 2 == 0 and cell.row > (3 if self.j_calib else 1):
                        cell.border = styles.Border(left=thin)
                row[0].alignment = styles.Alignment(horizontal='right')
                
            for i in range((len(list(ws.columns)) - 1) // (2*self.j_sets)):
                ws.cell(row=1, column=2*self.j_sets*i + 2).font = styles.Font(bold=True)
                ws.merge_cells(start_row=1, start_column=2 + i*2*self.j_sets,
                                 end_row=1,   end_column=1 + (i + 1)*2*self.j_sets)
                for j in range(self.j_sets):
                    ws.merge_cells(start_row=2, start_column=2*(self.j_sets*i + j + 1),
                                     end_row=2,   end_column=2*(self.j_sets*i + j + 1) + 1)
                if self.j_calib:
                    ws.merge_cells(start_row=3, start_column=4 + i*2*self.j_sets,
                                     end_row=3,   end_column=1 + (i + 1)*2*self.j_sets)
            self._basic_ws_format(ws, 4)

    def _export_table_views(self, wb, time_fmt_str):
        nobord = styles.Side(border_style='none', color='000000')
        thin = styles.Side(border_style='thin', color='000000')
        thick = styles.Side(border_style='thick', color='000000')

        ws_overall = wb.create_sheet("Competition Tables")
        ws_overall.append([val for pair in zip(2*self.t_pairs*[''], self.rooms[3][:2*self.t_pairs]) 
                   for val in pair])
        t_pair_sheets = [wb.create_sheet(room[:-2]) for room in self.rooms[3][:2*self.t_pairs:2]]
        for t_pair in range(self.t_pairs):
            t_pair_sheets[t_pair].append([''] + [self.rooms[3][2*t_pair]] 
                                       + [''] + [self.rooms[3][2*t_pair + 1]])

        for slot in self.t_slots:
            if slot is None:
                for ws in t_pair_sheets + [ws_overall]:
                    ws.append([''])
            else:
                line = sum([['None', ''] if t is None else [self.teams[t].num, self.teams[t].name]
                    for t in slot[2]], [slot[0].strftime(time_fmt_str)])
                ws_overall.append(line)
                for t_pair in range(self.t_pairs):
                    t_pair_sheets[t_pair].append([line[0]] + line[4*t_pair + 1: 4*t_pair + 5])

        for ws in [ws_overall] + t_pair_sheets:
            for row in ws.rows:
                for cell in row:
                    cell.alignment = styles.Alignment(horizontal='center')
                    if cell.column % 4 == 0:
                        cell.border = styles.Border(left=thin)
                    if cell.column % 4 == 2:
                        cell.border = styles.Border(left=thick)
                row[0].alignment = styles.Alignment(horizontal='right')
            self._basic_ws_format(ws)
    
    def _export_team_views(self, wb, time_fmt_str):
        ws_chron = wb.create_sheet("Team View (Chronological)")
        ws_event = wb.create_sheet("Team View (Event)")
        ws_chron.append(['Team Number', 'Team Name'] + ['Event {}'.format(i + 1) 
                            for i in range(len(self.teams[0].events))])
        ws_event.append(['Team Number', 'Team Name'] + self.event_names)
        
        for team in self.teams:
            ws_chron.append([team.num, team.name] + ['{} at {}, {}'.format(self.event_names[min(3, cat)], 
                time.strftime(time_fmt_str), self.rooms[min(3, cat)][loc]) 
                for (time, duration, cat, loc) in team.events])
            ws_event.append([team.num, team.name] + ['{}, {}'.format(time.strftime(time_fmt_str),
                self.rooms[min(3, cat)][loc]) for (time, duration, cat, loc) 
                in sorted(team.events, key=lambda x: x[2])])

        for ws in (ws_chron, ws_event):
            self._basic_ws_format(ws)

    def _basic_ws_format(self, ws, start=0):
        """bolds the top row, stripes the rows, and sets column widths"""
        for col in ws.columns:
            col[0].font = styles.Font(bold=True) 
            length = 1.2*max(len(str(cell.value)) for cell in col[start:])
            ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = length
        for row in list(ws.rows)[max(2, start - 1)::2]:
            for cell in row:
                cell.fill = styles.PatternFill('solid', fgColor='DDDDDD')
    
    def _team(self, team_num):
        return self.teams[team_num % self.num_teams]

    def read_data(self, fpath):
        dfs = pd.read_excel(fpath, sheet_name=["Team Information", "Input Form"])

        self.team_sheet = dfs["Team Information"]
        if all([x in self.team_sheet.columns for x in ("Team Number", "Team")]):
            self.teams = [Team(*x) for x in self.team_sheet.loc[:, ("Team Number", "Team")].values]
        else:
            raise KeyError("Could not find columns 'Team Number' and 'Team' in sheet 'Team Information'")

        param_sheet = dfs["Input Form"]
        if any([x not in param_sheet.columns for x in ("key", "answer")]):
            raise KeyError("Could not find columns 'key' and 'answer' in sheet 'Input Form'")
        self.param_sheet = param_sheet.set_index("key")
        param = dict(self.param_sheet.loc[:, "answer"].items())

        self.num_teams = len(self.teams)
        try:
            self.tournament_name = param["tournament_name"]
            self.scheduling_method = param["scheduling_method"]
            self.travel = timedelta(minutes=param["travel_time"])
            self.event_names = ['Project', 'Robot Design', 'Core Values']\
                             + self.param_sheet.loc["t_round_names"].dropna().values.tolist()[1:]
            
            self.j_start = datetime(1, 1, 1, param["j_start"].hour, param["j_start"].minute)
            self.j_sets = param["j_sets"]
            self.rooms = [self.param_sheet.loc[key].dropna().values.tolist()[1:]
                         for key in ("j_project_rooms", "j_robot_rooms", "j_values_rooms")]
            self.j_calib = (param["j_calib"] == "Yes")
            self.j_duration = timedelta(minutes=param["j_duration"])
            self.j_duration_team = timedelta(minutes=10)
            self.j_consec = param["j_consec"]
            self.j_break = timedelta(minutes=param["j_break"])
            self.j_lunch = datetime(1, 1, 1, param["j_lunch"].hour, param["j_lunch"].minute)
            self.j_lunch_duration = timedelta(minutes=param["j_lunch_duration"])
            
            self.t_pairs = param["t_pairs"]
            self.rooms += [sum([[tbl + ' A', tbl + ' B'] for tbl in self.param_sheet.loc["t_pair_names"]
                .dropna().values.tolist()[1:]], [])][:2*self.t_pairs]
            self.t_rounds = len(self.event_names[-1])
            self.t_duration = [timedelta(minutes=x) for x in 
                self.param_sheet.loc["t_durations"].dropna().values.tolist()[1:]]
            self.t_lunch = datetime(1, 1, 1, param["t_lunch"].hour, param["t_lunch"].minute)
            self.t_lunch_duration = timedelta(minutes=param["t_lunch_duration"])
            
        except KeyError as e:
            raise KeyError(str(e) + " not found in 'key' column in sheet 'Input Form'")
        
if __name__ == "__main__":
    Tournament().schedule()
