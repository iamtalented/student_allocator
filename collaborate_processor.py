import pprint

from openpyxl import load_workbook, Workbook

FILENAME = "Collaborate 3- Participant Sign-Up (Responses).xlsx"
FIRST_ROW = 2
NAME_COL = ["D", "E"]
SCHOOL_COL = "H"
EMAIL_COL = "J"
PHONE_COL = "K"
CHOICE_COL = "L"
CONSENT_COL = "P"

"""
Load the students and return a list of signups and a list
"""
def load_students():
    workbook = load_workbook(FILENAME)
    signups = []
    emails = []
    choices = {}
    for sheet_name in workbook.get_sheet_names():
        sheet = workbook[sheet_name]
        finished = False
        cur_row = FIRST_ROW
        while not finished:
            cur_row_str = str(cur_row)
            new_signup = {
                "name": sheet[NAME_COL[0] + cur_row_str].value + " " +
                        sheet[NAME_COL[1] + cur_row_str].value,
                "school": sheet[SCHOOL_COL + cur_row_str].value,
                "email": sheet[EMAIL_COL + cur_row_str].value,
                "phone": sheet[PHONE_COL + cur_row_str].value,
                "choices": [x.strip() for x in sheet[CHOICE_COL + cur_row_str].value.split(",")],
                "consent": sheet[CONSENT_COL + cur_row_str].value
            }
            cur_row += 1
            if sheet[NAME_COL[0] + str(cur_row + 1)].value is None:
                finished = True
            if new_signup["consent"] != "Yes" or new_signup["email"] in emails:
                continue
            emails.append(new_signup["email"])
            for choice in new_signup["choices"]:
                if choice == "Systems" or choice == "Engineering":
                    new_signup["choices"].remove(choice)
                    continue
                if choice == "Technology: Big Data":
                    new_signup["choices"].remove(choice)
                    choice = "Technology: Big Data, Systems, Engineering"
                    new_signup["choices"].append(choice)
                    continue
            for choice in new_signup["choices"]:
                if choice not in choices:
                    choices[choice] = 0
                choices[choice] += 1
            signups.append(new_signup)
    return signups, choices

"""
Filtered demand is a list of classes that have met the min threshold to get in
and is in order to lowest demand to highest
"""
def sort_students(students, filtered_demand):
    print students

if __name__ == "__main__":
    students, counts = load_students()
    counts = sorted(counts.iteritems(), key=lambda (k, v): (v, k))
    for line in counts:
        if line[1] > 15:
            print line[0] + ":"



