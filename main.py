import glob
import re

import pandas as pd
from openpyxl.utils import get_column_letter

studio_regex = re.compile(r'(ST(UDIOUL|\.)\.?|CAR\s|PANGRATTI)\s?(([^HAND])(.*))', flags=re.IGNORECASE)
time_regex = re.compile(r'(^\d.*\d)\s?([^0-9.\-:]+)?')
studios = {}

file = ''
sheet = ''

persons = ['badea', 'voislav']
job_type = 'Cameraman'
types = ['B', 'E', 'CT']

headers = ['Nume', 'Luna', 'Ziua', 'Structura', 'Functie', 'Program', 'Perioada', 'Tip', 'Filename', 'Sheet', 'Cell']
export_df = pd.DataFrame(columns=headers)


class Position:
    def __init__(self, filename, sheet, row, column):
        self.column = column
        self.row = row
        self.sheet = sheet
        self.filename = filename
        self.cell = get_column_letter(self.column + 1) + str(self.row)

    def __str__(self):
        return f'{self.filename} - {self.sheet} - {self.cell}'

    def __repr__(self):
        return self.__str__()


class Studio:
    def __init__(self, name, type, programs):
        self.programs = programs
        self.type = type
        self.name = name

    def __str__(self):
        return f'{self.name} - {self.type} - {self.programs}'

    def __repr__(self):
        return self.__str__()


class Program:
    def __init__(self, name, activities, people, source):
        self.source = source
        self.people = people
        self.activities = activities
        self.name = name

    def __str__(self):
        return f'{self.name} - {self.people} - {self.activities} - {self.source}'

    def __repr__(self):
        return self.__str__()


class Activity:
    def __init__(self, time, type):
        self.type = type
        self.time = time

    def __str__(self):
        return f'{self.time} - {self.type}'

    def __repr__(self):
        return self.__str__()


def is_studio(string):
    match = re.match(studio_regex, str(string))
    return match


def get_studio_list(df):
    global studios
    studios = {}
    for column in range(len(df.columns)):
        for cell in df[df.columns[column]]:
            match = is_studio(cell)
            cell = str(cell).lower()
            if match or cell == 'pangratti':
                studio_type = ''

                if cell == 'pangratti':
                    name = 'Pangratti'
                    studio_type = 'studio'
                else:
                    name = match.group(3)

                if 'st' in cell:
                    studio_type = 'studio'
                elif 'car' in cell:
                    studio_type = 'car'
                elif 'pangratti' not in cell:
                    studio_type = 'unknown'

                studios[name] = Studio(name, studio_type, [])
    studios['Unknown'] = Studio('Unknown', 'Unknown', [])


def get_next_filled_cell(df, row, column, direction=1, regex=r''):
    row += direction
    while (pd.isnull(df[df.columns[column]][row]) or (regex != r'' and not re.match(regex, df[df.columns[column]][row]))) and len(df) - 1 > row > 0:
        row += direction
    return row, column, df[df.columns[column]][row]


def get_program(df, time_row, column):
    if time_row == len(df) - 1:
        return len(df)
    activities = []
    names = []
    title = get_next_filled_cell(df, time_row, column, -1)
    title_name = title[2]
    studio = get_next_filled_cell(df, time_row, column, -1, studio_regex)
    times_range = (time_row, get_regex_until_ne(df, time_row, column, 1, time_regex))
    try:
        studio_name = is_studio(studio[2]).group(3) if is_studio(studio[2]).group(3) else studio[2]
    except:
        if studio[2] == 'PANGRATTI':
            studio_name = 'Pangratti'
        else:
            studio_name = 'Unknown'
    names_range = [times_range[1] + 1, get_regex_until_ne(df, times_range[1] + 1, column, 1, r'^\D+')]

    # if get_next_filled_cell(df, time_row, column, 1, studio_regex)[0] != len(df) - 1:
    #     names_range[1] = get_next_filled_cell(df, time_row, column, 1, studio_regex)[0] - 1
    # else:
    #     names_range[1] = get_next_filled_cell(df, time_row, column, -1)[0] - 1

    for row in range(times_range[0], times_range[1] + 1):
        if str(get_cell(df, row, column)) == 'nan':
            continue
        match = re.match(time_regex, get_cell(df, row, column))
        if match:
            time = match.group(1)
            type = match.group(2) if match.group(2) else 'B'
            activities.append(Activity(time, type))

    for row in range(names_range[0], names_range[1] + 1):
        cell = str(get_cell(df, row, column))
        if cell != 'nan' and not any(x in cell for x in ['ORE', 'INTIRZ', 'PROSP', 'FILM', 'NOAPTE', 'ATENTIE', 'TINUTA', 'LEGIT', 'BULETIN', 'PROGRAM', 'PLECARE', 'PULSUL', 'HANDBAL', 'VREMEA', 'PRIETEN', 'TURA', 'JURNAL', '\\']):
            name = re.match(r'^[a-zA-Z\s]+', get_cell(df, row, column).lower()).group(0).strip()
            names.append(name)
            # if name not in persons:
            #     persons.append(name)
    studios[studio_name].programs.append(Program(title_name if studios[studio_name].name != '11' else 'Stiri', activities, names, Position(file[9:], sheet, title[0] + 2, column)))

    return times_range[1] + 1


def get_program_list(df):
    for column in range(len(df.columns)):
        row = 0
        prev_row = -1
        while prev_row != row and row < len(df) - 1:
            prev_row = row
            row = get_program(df, get_next_filled_cell(df, row, column, 1, time_regex)[0], column)


def get_cell(df, row, column):
    return df[df.columns[column]][row]


def parse_sheet(df, first_name, last_name):
    global export_df

    get_studio_list(df)
    get_program_list(df)
    first_name = first_name.lower()
    last_name = last_name.lower()
    first_name_initial = first_name[0]
    date = re.match(r'(\w+)\s*(\d+).(\d+).(\d+)', df.columns[0])
    year = date.group(4)
    if year != '2018':
        year = '2018'
    month = date.group(3)
    day = date.group(2)
    day_name = date.group(1)
    date = f'{year}.{month}.{day} ({day_name.capitalize()})'
    month_names = {1: 'Ianuarie', 2: 'Februarie', 3: 'Martie', 4: 'Aprilie', 5: 'Mai', 6: 'Iunie', 7: 'Iulie', 8: 'August', 9: 'Septembrie', 10: 'Octombrie', 11: 'Noiembrie', 12: 'Decembrie'}

    possible_names = [first_name]
    if last_name:
        last_name_initial = last_name[0]
        possible_names.append([last_name + ' ' + first_name, last_name + ' ' + first_name_initial, first_name + ' ' + last_name, first_name + ' ' + last_name_initial])

    for key, studio in studios.items():
        studio_name = f'{studio.type.capitalize()} {studio.name.capitalize()}'
        for program in studio.programs:
            if any(name in program.people for name in possible_names):
                for activity in program.activities:
                    export_df = export_df.append({'Luna': month_names[int(month)], 'Ziua': date, 'Structura': studio_name, 'Functie': job_type, 'Program': program.name.title(), 'Perioada': activity.time, 'Tip': activity.type, 'Filename': program.source.filename, 'Sheet': program.source.sheet, 'Cell': program.source.cell}, ignore_index=True)
    # for key, studio in studios.items():
    #     print(studio)


def get_regex_until_ne(df, row, column, direction, regex_to_find):
    while get_next_filled_cell(df, row, column, direction, regex_to_find)[0] < len(df) - 1:
        if not re.match(regex_to_find, get_next_filled_cell(df, row, column, direction)[2]):
            return row
        row = get_next_filled_cell(df, row, column, direction, regex_to_find)[0]
    return row


def parse_file(filename, first_name, last_name):
    global file, sheet
    file = filename
    sheets = pd.read_excel(filename, sheet_name=None)
    sheet_names = list(sheets.keys())[1:]
    for sheet_name in sheet_names:
        sheet = sheet_name
        df = sheets[sheet_name]
        parse_sheet(df, first_name, last_name)


def export(first_name, last_name):
    global export_df
    full_name = f'{first_name} {last_name}'.title().strip()
    dfs = []

    for type in types:
        dfs.append(export_df[export_df['Tip'].str.contains(type)])
    export_df = pd.concat(dfs)
    export_df.sort_values(['Ziua', 'Perioada'], ascending=True, inplace=True)
    export_df.reset_index(inplace=True, drop=True)
    export_df['Nume'][0] = full_name

    export_df.to_excel(f'Raport - {full_name}.xlsx', '2018', index=False)
    export_df.to_excel(f'\\\\192.168.0.178\\homes\\Vali\\Raport - {full_name}.xlsx', '2018', index=False)
    print(f'Created raport for {full_name}')


def parse_folder(first_name, last_name=''):
    files = glob.glob('programs\\*.xls')
    for file in files:
        parse_file(file, first_name, last_name)

    export(first_name, last_name)


for person in persons:
    parse_folder(person)
