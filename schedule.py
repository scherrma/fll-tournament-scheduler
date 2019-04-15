#!/usr/bin/env python3
"""Module for use in scheduling one-day FLL tournaments."""
from datetime import datetime, timedelta
import os
import sys
import tkinter
from tkinter import filedialog
import pandas
import openpyxl
import openpyxl.styles as styles
from openpyxl.utils import get_column_letter
from scheduler.tournament import Tournament

def read_data(fpath):
    """Imports the team roster and scheduling settings from the input form."""
    dfs = pandas.read_excel(fpath, sheet_name=["Team Information", "Input Form"])

    team_sheet = dfs["Team Information"]
    column_check = ["Team Number", "Team"]
    if all([x in team_sheet.columns for x in column_check]):
        divisions = ("Division" in team_sheet.columns)
        if divisions:
            column_check += ["Division"]
        teams = team_sheet.loc[:, column_check].values
        team_info_base = [(i, list(team_sheet.columns).index(cat) + 1) for i, cat in
                          list(enumerate(column_check[::-1]))]
        team_info = ["=index(indirect(\"'Team Information'!C{1}\", false),"\
                      "match(indirect(\"RC[-{2}]\", false), "\
                      "indirect(\"'Team Information'!C{0}\", false), 0))"
                     .format(team_info_base[-1][1], cat, offset + 1) for offset, cat
                     in team_info_base[:-1]]
    else:
        raise KeyError("Could not find columns 'Team Number' and 'Team' in 'Team Information'")

    param_sheet = dfs["Input Form"]
    if any([x not in param_sheet.columns for x in ("key", "answer")]):
        raise KeyError("Could not find columns 'key' and 'answer' in sheet 'Input Form'")
    param_sheet = param_sheet.set_index("key")
    param = dict(param_sheet.loc[:, "answer"].items())

    try:
        tournament_name = param["tournament_name"]
        scheduling_method = param["scheduling_method"]
        travel = timedelta(minutes=param["travel_time"])
        event_names = ["Coaches' Meeting", "Opening Ceremonies", 'Project', 'Robot Design',
                       'Core Values']
        event_names += param_sheet.loc["t_round_names"].dropna().values.tolist()[1:]
        t_rounds = len(event_names) - 5
        coach_meet = (datetime.combine(datetime(1, 1, 1), param["coach_start"]),
                      timedelta(minutes=param["coach_duration"]))
        opening = (datetime.combine(datetime(1, 1, 1), param["opening_start"]),
                   timedelta(minutes=param["coach_duration"]))

        j_start = datetime.combine(datetime(1, 1, 1), param["j_start"])
        j_sets = param["j_sets"]
        rooms = [[param["coach_room"]], [param["opening_room"]]]
        rooms += [param_sheet.loc[key].dropna().values.tolist()[1:]
                  for key in ("j_project_rooms", "j_robot_rooms", "j_values_rooms")]
        j_calib = (param["j_calib"] == "Yes")
        j_duration = (timedelta(minutes=param["j_duration"]), timedelta(minutes=10))
        j_breaks = (param["j_consec"], timedelta(minutes=param["j_break"]))
        j_lunch = (datetime.combine(datetime(1, 1, 1), param["j_lunch"]),
                   timedelta(minutes=param["j_lunch_duration"]))

        t_pairs = param["t_pairs"]
        rooms += [sum([[tbl + ' A', tbl + ' B'] for tbl in
                       param_sheet.loc["t_pair_names"].dropna().values.tolist()[1:]],
                      [])][:2*t_pairs]
        t_duration = [timedelta(minutes=x) for x in
                      param_sheet.loc["t_durations"].dropna().values.tolist()[1:]]
        t_lunch = (datetime.combine(datetime(1, 1, 1), param["t_lunch"]),
                   timedelta(minutes=param["t_lunch_duration"]))

    except KeyError as excep:
        raise KeyError(str(excep) + " not found in 'key' column in sheet 'Input Form'")

    return ((teams, divisions, scheduling_method, travel, coach_meet, opening, j_start, j_sets,
             j_calib, j_duration, j_breaks, j_lunch, t_rounds, t_pairs, t_duration, t_lunch),
            tournament_name, (team_info, event_names, rooms))

def export(tment, workbook, team_info, event_names, rooms):
    """Exports schedule to an xlsx file; uses the tournament name for the file name."""
    for sheet in [ws for ws in workbook.sheetnames if ws != 'Team Information']:
        del workbook[sheet]

    time_fmt = "%{}I:%M %p".format('#' if sys.platform == "win32" else '-')
    export_judge_views(tment, workbook, time_fmt, team_info, event_names, rooms)
    export_table_views(tment, workbook, time_fmt, team_info, rooms)
    export_team_views(tment, workbook, time_fmt, team_info, event_names, rooms)

def export_judge_views(tment, workbook, time_fmt, team_info, event_names, rooms):
    """Adds the four judging-focused sheets to the output workbook."""
    thin = styles.Side(border_style='thin', color='000000')
    thick = styles.Side(border_style='thick', color='000000')
    team_width = 1 + len(team_info)

    sheets = [workbook.create_sheet(name) for name in ["Judging Rooms"] + event_names[2:5]]

    rows = [[['']], [['']]]
    for i in range(3):
        rows[0].append([event_names[i + 2]] + (tment.j_sets*team_width - 1)*[''])
        rows[1].append(sum([[room] + (team_width - 1)*[''] for room in rooms[i + 2]], []))

    for time, teams in tment.j_slots:
        rows.append([[time.strftime(time_fmt)]])
        if teams is not None:
            teams = [[None if t is None else tment.teams[t] for t in cat] for cat in teams]
            if len(teams[0]) == 1 and tment.j_calib:
                rows[-1] += [[teams[i][0].num] + team_info
                             + ["all {} judges in {}".format(event_names[i + 2].lower(),
                                                             rooms[i + 2][0])]
                             + (team_width*(tment.j_sets - 1) - 1)*[''] for i in range(3)]
            else:
                rows[-1] += [sum([['']*(team_width - 1) + ['None'] if team is None else
                                  [team.num] + team_info for team in cat], []) for cat in teams]
    for row in rows:
        sheets[0].append(sum(row, []))
        for i in range(3):
            sheets[i + 1].append(row[0] + (row[i + 1] if len(row) > 1 else []))

    col_sizes = [1 + max(len(str(text)) for text in cat) for cat in
                 zip(*[team.info(tment.divisions) for team in tment.teams])]
    for sheet in sheets:
        basic_sheet_format(sheet, 4)
        for col in list(sheet.columns)[1:]:
            sheet.column_dimensions[get_column_letter(col[0].column)].width =\
                    col_sizes[(col[0].column - 2) % len(col_sizes)]
        sheet_borders(sheet, ((styles.Border(left=thick), team_width*tment.j_sets, 2, 0),
                              (styles.Border(left=thin), team_width, 2, 1 + 2*tment.j_calib)))

        for i in range((len(list(sheet.columns)) - 1) // (team_width*tment.j_sets)):
            sheet.cell(row=1, column=team_width*tment.j_sets*i + 2).font = styles.Font(bold=True)
            sheet.merge_cells(start_row=1, start_column=2 + i*team_width*tment.j_sets,
                              end_row=1, end_column=1 + (i + 1)*team_width*tment.j_sets)
            for j in range(tment.j_sets):
                sheet.merge_cells(start_row=2, start_column=team_width*(tment.j_sets*i + j) + 2,
                                  end_row=2, end_column=team_width*(tment.j_sets*i + j + 1) + 1)
            if tment.j_calib:
                sheet.merge_cells(start_row=3, start_column=2 + team_width*(tment.j_sets*i + 1),
                                  end_row=3, end_column=1 + team_width*(tment.j_sets*(i + 1)))

def export_table_views(tment, workbook, time_fmt, team_info, rooms):
    """Adds the competition table focused sheets to the output workbook."""
    thin = styles.Side(border_style='thin', color='000000')
    thick = styles.Side(border_style='thick', color='000000')
    team_width = 1 + len(team_info)

    sheet_overall = workbook.create_sheet("Competition Tables")
    header = sum([[tbl] + (team_width - 1)*[''] for tbl in rooms[5]], [''])
    sheet_overall.append(header)
    t_pair_sheets = [workbook.create_sheet(room[:-2]) for room in rooms[5][::2]]
    for t_pair in range(tment.t_pairs):
        t_pair_sheets[t_pair].append([''] + header[2*team_width*t_pair + 1:
                                                   2*team_width*(t_pair + 1)])

    for slot in tment.t_slots:
        if slot is None:
            for sheet in t_pair_sheets + [sheet_overall]:
                sheet.append([''])
        elif all([team is None for team in slot[2]]):
            for sheet in t_pair_sheets + [sheet_overall]:
                sheet.append([slot[0].strftime(time_fmt)])
        else:
            line = sum([(team_width - 1)*[''] + ['None'] if t is None else
                        [tment.teams[t].num] + team_info for t in slot[2]],
                       [slot[0].strftime(time_fmt)])
            sheet_overall.append(line)
            for t_pair in range(tment.t_pairs):
                t_pair_sheets[t_pair].append([line[0]] + line[2*team_width*t_pair + 1:
                                                              2*team_width*(t_pair + 1) + 1])

    col_sizes = [1 + max(len(str(text)) for text in cat) for cat in
                 zip(*[team.info(tment.divisions) for team in tment.teams])]
    for sheet in [sheet_overall] + t_pair_sheets:
        basic_sheet_format(sheet, 2)
        for col in sheet.columns:
            if col[0].column > 1:
                sheet.column_dimensions[get_column_letter(col[0].column)].width =\
                        col_sizes[(col[0].column - 2) % len(col_sizes)]
        sheet_borders(sheet, ((styles.Border(left=thick), 2*team_width, 2, 0),
                              (styles.Border(left=thin), 2*team_width, 2 + team_width, 0)))
        for i in range(2, len(list(sheet.columns)), team_width):
            sheet.merge_cells(start_row=1, start_column=i,
                              end_row=1, end_column=i + team_width - 1)

def export_team_views(tment, workbook, time_fmt, team_info, event_names, rooms):
    """Adds event-sorted and time-sorted team-focused views to the output workbook."""
    ws_chron = workbook.create_sheet("Team View (Chronological)")
    ws_event = workbook.create_sheet("Team View (Event)")
    team_header = ['Team Number'] + (['Division'] if tment.divisions else []) + ['Team Name']
    ws_chron.append(team_header + ['Event ' + str(i + 1) for i in range(len(event_names))])
    ws_event.append(team_header + event_names[2:])

    for team in sorted(tment.teams, key=lambda t: t.num):
        ws_chron.append([team.num] + team_info
                        + ['{} at {}, {}'.format(event_names[cat], time.strftime(time_fmt),
                                                 rooms[min(5, cat)][loc])
                           for (time, duration, cat, loc) in team.events])
        ws_event.append([team.num] + team_info
                        + ['{}, {}'.format(time.strftime(time_fmt), rooms[min(5, cat)][loc])
                           for (time, length, cat, loc)
                           in sorted(team.events, key=lambda x: x[2])[2:]])

    col_size = 1 + max(len(team.name) for team in tment.teams)
    for sheet in (ws_chron, ws_event):
        basic_sheet_format(sheet)
        for col in sheet.columns:
            if col[0].column == 2 + tment.divisions:
                sheet.column_dimensions[get_column_letter(col[0].column)].width = col_size

def basic_sheet_format(sheet, start=0):
    """bolds the top row, stripes the rows, and sets column widths"""
    for col in sheet.columns:
        col[0].font = openpyxl.styles.Font(bold=True)
        length = 2 + max((len(str(cell.value)) for cell in col[start:]
                          if cell.value and str(cell.value)[0] != '='), default=0)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = length
    for row in list(sheet.rows):
        row[0].alignment = openpyxl.styles.Alignment(horizontal='right')
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            if cell.row > 1 and cell.row % 2 == 0:
                cell.fill = openpyxl.styles.PatternFill('solid', fgColor='DDDDDD')

def sheet_borders(sheet, borders=()):
    """Formats worksheet to add specified borders at a given period and offset."""
    for row in sheet.rows:
        for cell in row:
            for border, period, offset, start_row in borders[::-1]:
                if (cell.column - offset) % period == 0 and cell.row > start_row:
                    cell.border = border

def generate_schedule():
    """Top-most level function; gets a file, reads and schedules for it, then exports the result."""
    if len(sys.argv) == 1:
        root = tkinter.Tk()
        root.withdraw()
        fpath = filedialog.askopenfilename(initialdir=os.path.dirname(os.path.abspath(__file__)),
                                           filetypes=[("Excel files", "*.xls *.xlsm *.xlsx")])
        root.destroy()
    elif len(sys.argv) == 2:
        fpath = sys.argv[1]
    else:
        raise SystemExit("{} only accepts either zero or one files as input".format(sys.argv[0]))

    try:
        logic_params, tournament_name, io_params = read_data(fpath)
        tment = Tournament(*logic_params)
        tment.schedule()

        workbook = openpyxl.load_workbook(fpath)
        export(tment, workbook, *io_params)
        outfpath = os.path.join(os.path.dirname(fpath),
                                tournament_name.lower().replace(' ', '_') + '_schedule{}.xlsx')
        saved, attempts = False, 0
        while not saved:
            try:
                final_fout = outfpath.format(' ({})'.format(attempts) if attempts else '')
                workbook.save(final_fout)
                saved = True
            except PermissionError:
                attempts += 1
        print('Schedule saved: {}'.format(final_fout))

    except Exception as excep:
        raise excep
        #print(excep)

    if sys.platform in ["win32", "darwin"]:
        os.system("pause")

if __name__ == "__main__":
    print("Python environment initialized")
    generate_schedule()
