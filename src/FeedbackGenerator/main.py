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

    template_root = r"./templates"
    data_root = r"./data"

    templates = os.listdir(template_root)
    student_files = os.listdir(data_root)

    for exercise in templates:
        for groups in student_files:
            exercise_file = open(f'{template_root}/{exercise}', "r");
            students_file = open(f'{data_root}/{groups}', "r");

            tasks = pd.read_csv(filepath_or_buffer=f'{template_root}/{exercise}', skiprows=4, delimiter=";");
            students = pd.read_csv(filepath_or_buffer=f'{data_root}/{groups}', skiprows=1);

            assignment = Assignment();

            assignment.author = get_next_element(exercise_file);
            assignment.lecture = get_next_element(exercise_file);
            assignment.excercise = get_next_element(exercise_file);

            group = get_next_element(students_file);
            file_name = f"{assignment.lecture}_{assignment.excercise}_G{group}";
            f_name = f'generated/';

            isExist = os.path.exists(f'{f_name}')

            if not isExist:
                # Create a new directory because it does not exist
                os.makedirs(f'{f_name}')

            workbook = xlsxwriter.Workbook(f'./{f_name}/{file_name}.xlsx')
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
                start = 3;

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


def create_solution_file(fname):
    split = fname.split("_")

    group_name = split[len(split) - 1].lower()

    filename = f"./generated/{fname}.xlsx"
    stud_file = f"./data/students_{group_name}.csv"

    students = pd.read_csv(filepath_or_buffer=stud_file, skiprows=1)

    isExist = os.path.exists(f'./grading/{fname}')

    if not isExist:
        # Create a new directory because it does not exist
        os.makedirs(f'./grading/{fname}')

    for name in students["name"]:
        df = pd.read_excel(filename, sheet_name=name, index_col=[0])
        df.rename( columns={'Unnamed: 4':''}, inplace=True )
        df = df.replace(np.nan, '', regex=True)
        df.to_html(f"./grading/{fname}/{name}.html")

        file = open(filename, 'r', encoding="utf8", errors='replace').read()
        file.replace("<table ", "<table class=\"" + table_class + "\" ")

# main.py
import sys

if __name__ == "__main__":
    if sys.argv[1] == "-excel":
        create_grading_file()

    if sys.argv[1] == "-solution":
        if not sys.argv[2] == "":
            create_solution_file(sys.argv[2])