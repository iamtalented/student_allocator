import iat_2023_dal

import json
from pprint import pprint
import random

CLASS_LIMITS_FILE = "./data/quota.json"
MAX = 20
SESSION_COUNT = 3
STUDENT_CHOICES = 5
STUDENT_MAPPING_FILE = "./data/mapping.json"

def sort_students(signup_list):
    random.shuffle(signup_list)
    classes = generate_class_demand(signup_list)
    unassigned_students = []
    for student in signup_list:
        class_priority = generate_priority_order(student, classes)
        unassigned = False
        for i in range(SESSION_COUNT):
            session_id = "session"+str(i+1)
            assigned = False
            for _class in class_priority:
                if classes[_class][session_id] >= classes[_class]["max"]:
                    continue
                else:
                    student[session_id] = _class
                    classes[_class][session_id] += 1
                    class_priority.remove(_class)
                    assigned = True
                    break
            if not assigned:
                student[session_id] = None
                unassigned = True
        if unassigned:
            unassigned_students.append(student)
    #print("Unassigned Students")
    #print(unassigned_students)
    return signup_list, unassigned_students

            

def generate_class_demand(signup_list):
    with open(CLASS_LIMITS_FILE) as classes_json_file:
        classes = json.load(classes_json_file)
        for _class in classes:
            classes[_class]["demand"] = 0
            classes[_class]["session1"] = 0
            classes[_class]["session2"] = 0
            classes[_class]["session3"] = 0
        for student in signup_list:
            for choice in student["choices"]:
                if choice in classes and "demand" in classes[choice]:
                    classes[choice]["demand"] += 1
                else:
                    raise(choice + " not in list of classes")
        return classes

def generate_priority_order(student, classes):
    class_priority_scores = []
    to_fill = 0
    if len(student["choices"]) < STUDENT_CHOICES:
        to_fill = STUDENT_CHOICES - len(student["choices"])
        class_signup_counts = []
        for _class in classes:
            class_signup_counts.append([_class, classes[_class]["demand"]])
        class_signup_counts = sorted(class_signup_counts, key=lambda _class: _class[1])
        #print(class_signup_counts)
        base_counter = 0
        for i in range(to_fill):
            while class_signup_counts[base_counter + i][0] in student["choices"]:
                base_counter += 1
            student["choices"].append(class_signup_counts[base_counter + i][0])
            classes[class_signup_counts[base_counter + i][0]]["demand"] += 1
    for _class in student["choices"]:
        max = classes[_class]["max"]
        signup_coefficient = (1.0*classes[_class]["session1"]/max + 1.0*classes[_class]["session2"]/max + 1.0*classes[_class]["session3"]/max) / 3
        demand_coefficient = (1.0*classes[_class]["demand"]/(3*classes[_class]["max"]))
        score =  demand_coefficient - signup_coefficient
        class_priority_scores.append((_class, score))
        #print({"name": _class, "s1": signup_coefficient, "s2":demand_coefficient, "s3": score})
    class_priority_scores = sorted(class_priority_scores, key=lambda score: -score[1])
    #print(class_priority_scores)
    return [_class[0] for _class in class_priority_scores]

def generate_id_student_mapping(signup_list):
    mapping = {}
    for student in signup_list:
        mapping[student["id"]] = [student["name"], student["school"]]
    with open(STUDENT_MAPPING_FILE, "w") as mapping_json:
        json.dump(mapping, mapping_json, indent=4, sort_keys=True)

if __name__ == '__main__':
    students = iat_2023_dal.load_students_tasters()
    class_list, unassigned = sort_students(students)
    id_student_list = generate_id_student_mapping(class_list)
    iat_2023_dal.export_student_tasters_to_excel(class_list)
