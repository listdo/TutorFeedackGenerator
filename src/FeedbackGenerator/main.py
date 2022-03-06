import pandas as pd
import numpy as np
import xlsxwriter
import os

table_class = "rwd-table"

html = """<style> table {border-collapse: collapse; width: 100%; } th, td {text-align: left; padding: 8px; } tr:nth-child(even) {background-color: #f2f2f2;} </style>"""
html += """
<script>
  window.console = window.console || function(t) {};
</script>
<script>
  if (document.location.search.match(/type=embed/gi)) {
    window.parent.postMessage("resize", "*");
  }
</script>"""


class Assignment:
    author = ""
    lecture = ""
    excercise = ""

def get_next_element(f):
    return f.readline().split(":")[1].rstrip();

def create_grading_file():
    exercise = r"./templates/cgb_ex1.temp"
    stud_file = r"./data/students_g1.csv"

    exercise_file = open(exercise, "r");
    students_file = open(stud_file, "r");

    tasks = pd.read_csv(filepath_or_buffer=exercise, skiprows=4, delimiter=";");
    students = pd.read_csv(filepath_or_buffer=stud_file, skiprows=1);

    assignment = Assignment();

    assignment.author = get_next_element(exercise_file);
    assignment.lecture = get_next_element(exercise_file);
    assignment.excercise = get_next_element(exercise_file);

    group = get_next_element(students_file);
    f_name = f'{assignment.lecture}_{assignment.excercise}_G{group}';

    isExist = os.path.exists(f'{f_name}')

    if not isExist:
        # Create a new directory because it does not exist
        os.makedirs(f'{f_name}')

    workbook = xlsxwriter.Workbook(f'./{f_name}/{f_name}.xlsx')
    bold = workbook.add_format({'bold': True})

    for name in students["name"]:
        worksheet = workbook.add_worksheet(name=name)

        worksheet.write(0, 0, f"Author: {assignment.author}");
        worksheet.write(0, 1, f"Student: {name}");
        worksheet.write(0, 2, f'Lecture: {assignment.lecture}');
        worksheet.write(0, 3, f'Assignment: {assignment.excercise}');


        worksheet.write(1, 0, f'Tasks', bold);
        worksheet.write(1, 1, f'Description', bold);
        worksheet.write(1, 2, f'Possible Points', bold);
        worksheet.write(1, 3, f'Reached Points', bold);
        worksheet.write(1, 4, f'Comment', bold);

        i = 2;
        start = 2;

        for (index, row) in tasks.iterrows():
            worksheet.write(i, 0, row["Task"])
            worksheet.write(i, 1, row["SubTask"])
            worksheet.write(i, 2, row["Points"])
            i += 1

        worksheet.write(i, 1, f'SUM', bold)
        worksheet.write(i, 2, f'=SUM(C{start}:C{i})')
        worksheet.write(i, 3, f'=SUM(D{start}:D{i})')

        for column in tasks:
            column_width = max(tasks[column].astype(str).map(len).max(), len(column))
            col_idx = tasks.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx + 10, column_width)

    workbook.close()

def create_solution_file():
    fname = "CGB4_BB_EX1_G1"

    filename = f"./{fname}/{fname}.xlsx"
    stud_file = r"./data/students_g1.csv"

    students = pd.read_csv(filepath_or_buffer=stud_file, skiprows=1)
    students_file = open(stud_file, "r")

    group = get_next_element(students_file)
    isExist = os.path.exists(f'{fname}')

    if not isExist:
        # Create a new directory because it does not exist
        os.makedirs(f'{fname}')

    for name in students["name"]:
        df = pd.read_excel(filename, sheet_name=name, index_col=[0])
        df = df.replace(np.nan, '', regex=True)
        df.to_html(f"./{fname}/{name}.html")

        print(df)

        file = open(filename, 'r', encoding="utf8", errors='replace').read()

        file = file.replace("<table ", "<table class=\"" + table_class + "\" ")

#create_grading_file()
create_solution_file()