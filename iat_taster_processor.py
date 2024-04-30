import iat_2023_dal

from pprint import pprint
import random

MAX = 20

def sort_students(signup_list):
    random.shuffle(signup_list)
    classes = generate_class_demand(signup_list)
    unassigned_students = []
    for student in signup_list:
        class_priority = generate_priority_order(student, classes)
        unassigned = False
        for i in range(3):
            session_id = "session"+str(i+1)
            assigned = False
            for _class in class_priority:
                if classes[_class][session_id] >= MAX:
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
    return signup_list, unassigned_students

            

def generate_class_demand(signup_list):
    classes = {}
    for student in signup_list:
        for choice in student["choices"]:
            if choice in classes:
                classes[choice]["demand"] += 1
            else:
                classes[choice] = {}
                classes[choice]["demand"] = 1
                classes[choice]["session1"] = 0
                classes[choice]["session2"] = 0
                classes[choice]["session3"] = 0
    return classes

def generate_priority_order(student, classes):
    class_priority_scores = []
    class_type_count = len(classes)
    total_selections = sum([classes[_class]["demand"] for _class in classes])
    for _class in student["choices"]:
        signup_coefficient = (1.0*classes[_class]["session1"]/MAX + 1.0*classes[_class]["session2"]/MAX + 1.0*classes[_class]["session3"]/MAX) / 3
        demand_coefficient = (1.0*classes[_class]["demand"]/total_selections) * class_type_count
        score = signup_coefficient - demand_coefficient
        class_priority_scores.append((_class, score))
        #print({"name": _class, "s1": signup_coefficient, "s2":demand_coefficient, "s3": score})
    class_priority_scores = sorted(class_priority_scores, key=lambda score: score[1])
    return [_class[0] for _class in class_priority_scores]

if __name__ == '__main__':
    students = iat_2023_dal.load_students_tasters()
    class_list, unassigned = sort_students(students)
    iat_2023_dal.export_student_tasters_to_excel(class_list)
