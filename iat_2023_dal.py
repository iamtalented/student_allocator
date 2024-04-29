import os
import pprint

from openpyxl import load_workbook, Workbook

FIRST_CHOICE_COL = "E"
FIRST_ROW = 6
NAME_COL = "B"
NUM_CLASSES = 10
SCHOOL_COL = "D"
SHEETNAME = "Reg List - IAT2024"
TITLE_ROW = "4"

def load_students_tasters():
    filenames = [filename for filename in os.listdir() if 'xlsx' in filename]
    students = []
    for filename in filenames:
        students.extend(load_students_from_excel_tasters(filename))
    return students

def load_students_from_excel_tasters(filename):
    workbook = load_workbook(filename)
    sheet = workbook[SHEETNAME]
    signups = []
    cur_row = FIRST_ROW
    finished = False
    while not finished:
        cur_row_str = str(cur_row)
        if sheet[NAME_COL[0] + cur_row_str].value is None:
            finished = True
            break
        new_student = {
            "name": sheet[NAME_COL + cur_row_str].value.strip(),
            "school": sheet[SCHOOL_COL + cur_row_str].value.strip(),
            "choices": [],
            "assigned": []
        }
        cur_row += 1
        for i in range(NUM_CLASSES):
            new_col = chr(ord(FIRST_CHOICE_COL) + i)
            if sheet[new_col + cur_row_str].value:
                class_selection = sheet[new_col + TITLE_ROW].value.replace("\n", " ")
                new_student['choices'].append(class_selection)
        if new_student['choices'] == []:
            no_class_count += 1
        else:
            signups.append(new_student)

    return signups

def export_student_tasters_to_excel():
    return


