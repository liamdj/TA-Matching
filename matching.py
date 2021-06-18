import csv
import argparse
from typing import Dict, List, Tuple
import numpy as np
from scipy import stats
from ortools.graph import pywrapgraph

DIGITS = 3

def construct_flow(student_prefs, prof_prefs, course_info):

    N, M = len(student_prefs.keys()), sum([target for target, _ in course_info.values()])
    source, sink = N + M, N + M + 1
    flow = pywrapgraph.SimpleMinCostFlow(2 + N + M)

    # Each student cannot fill >1 course slot
    for i in range(N):
        flow.AddArcWithCapacityAndUnitCost(source, i, 1, 0)

    # Each course slot cannot have >1 TA
    # Value of filling slot is reciprocal with slot index
    node = N
    for target, value in course_info.values():
        total = sum(1/i for i in range(1, target + 1))
        for i in range(1, target + 1):
            cost = -value / i / total
            flow.AddArcWithCapacityAndUnitCost(node, sink, 1, int(cost * 10 ** DIGITS))
            node += 1

    # Edge weights given by preferences
    for i, student in enumerate(student_prefs.keys()):
        node = N
        for course, info in course_info.items():
            target, _ = info
            if course in student_prefs[student] and student in prof_prefs[course]:
                cost = -(student_prefs[student][course] + prof_prefs[course][student]) 
                for slot in range(target):
                    flow.AddArcWithCapacityAndUnitCost(i, node + slot, 1, int(cost * 10 ** DIGITS))
            node += target       

    # Every student must TA a course
    flow.SetNodeSupply(source, N)
    flow.SetNodeSupply(sink, -N)

    return flow

def write_matching(filename: str, flow, students, courses, fixed_matches):
    with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Student", "Course", "Weight"])
            
            for arc in range(flow.NumArcs()):
                si, ci = flow.Tail(arc), flow.Head(arc) - len(students)
                if flow.Flow(arc) > 0 and si < len(students):
                    writer.writerow([students[si], courses[ci], -flow.UnitCost(arc) / 10 ** DIGITS])

            for s, c, _ in fixed_matches:
                writer.writerow([s, c, "Fixed"])


scale = { "Excellent": 3, "Good": 2, "Okay": 1, "Poor": 0 }

ADVISOR_WEIGHT = 1
REPEAT_WEIGHT = 1
FAVORITE_WEIGHT = 0.1
FILL_COURSE_WEIGHT = 10

def read_partial_matching(filename: str) -> List[Tuple[str, str, float]]:
    matchings = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            weight = float(row["Weight"]) if "Weight" in row else 0
            matchings.append((row["Student"], row["Course"], weight))
    
    return matchings

def read_student_prefs(filename: str) -> Dict[str, Dict[str, float]]:
    prefs = {}
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        courses = reader.fieldnames[5:-1] # will break when form changes
        for row in reader:
            options = {c: scale[row[c]] for c in courses if row[c] != "Unqualified"}
            scores = stats.zscore(list(options.values()))
            np.nan_to_num(scores, copy=False)
            options_scores = {c: s for c, s in zip(options.keys(), scores)}

            advisors = row["Advisor's Course"] 
            if advisors in options and options[advisors] >= 2:
                options_scores[advisors] += ADVISOR_WEIGHT
            repeated = row["Previous TA Experience"]
            if repeated in options and options[repeated] >= 2:
                options_scores[repeated] += REPEAT_WEIGHT
            favorite = row["Favorite Course"]
            if favorite in options and options[favorite] >= 2:
                options_scores[favorite] += FAVORITE_WEIGHT

            netid, _, _ = row["Username"].partition('@')
            prefs[row["Name"]] = options_scores

    return prefs


def read_prof_prefs(filename: str, students: List[str], courses: List[str]) -> Dict[str, Dict[str, float]]:
    prefs = {}
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        max_students = (len(reader.fieldnames) - 2) // 3
        for row in reader:
            # default student rating is 'okay'
            options = dict.fromkeys(students, 1)
            for i in range(1, max_students + 1):
                student = row["Student ({})".format(i)]
                rating = row["Rating ({})".format(i)]
                if student in options:
                    if rating == "Disaster":
                        del options[student]
                    elif rating != "":
                        options[student] = scale[rating]
            
            scores = stats.zscore(list(options.values()))
            np.nan_to_num(scores, copy=False)
            options_scores = {t: s for t, s in zip(options.keys(), scores)}
            prefs[row["Course"]] = options_scores

    for c in courses:
        if c not in prefs:
            prefs[c] = dict.fromkeys(students, 1)
    return prefs
            


def read_course_info(filename: str) -> Dict[str, Tuple[int, float]]:
    infos = {}
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            try:
                weight = FILL_COURSE_WEIGHT * float(row["Fill Weight"])
            except (ValueError, KeyError):
                weight = FILL_COURSE_WEIGHT
            infos[row["Course"]] = int(row["Target TAs"]), weight
    
    return infos

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Generate matching from input data.")
    parser.add_argument("student", metavar="STUDENT INPUT", help="csv file with student responses")
    parser.add_argument("prof", metavar="PROFESSOR INPUT", help="csv file with professor responses")
    parser.add_argument("course", metavar="COURSE INPUT", help="csv file with course information")
    parser.add_argument("--fixed", metavar="FIXED INPUT", help="csv file with required student-course matchings")
    parser.add_argument("--adjusted", metavar="ADJUSTED INPUT", help="csv file with adjustment weights for student-course matchings")
    parser.add_argument("--matchings", metavar="MATCHING OUTPUT", nargs='?', default='matchings.csv', help="location to write matching output")
    parser.add_argument("--additional", metavar="ADDITIONAL OUTPUT", nargs='?', default='add_TA.csv', help="location to write effects of gaining an additional TA for each course")
    parser.add_argument("--removal", metavar="REMOVAL OUTPUT", nargs='?', default='remove_TA.csv', help="location to write effects of removing each TA")
    args = parser.parse_args()
    
    student_prefs = read_student_prefs(args.student)
    course_info = read_course_info(args.course)

    if args.fixed:
        fixed_matches = read_partial_matching(args.fixed)
        for student, course, _ in fixed_matches:
            del student_prefs[student]
            if course_info[course][0] <= 1:
                del course_info[course]
            else:
                course_info[course] = course_info[course][0] - 1, course_info[course][1]
    else:
        fixed_matches = []

    students = list(student_prefs.keys())
    prof_prefs = read_prof_prefs(args.prof, students, list(course_info.keys()))

    if args.adjusted:
        adjusted_matches = read_partial_matching(args.adjusted)
        for student, course, weight in adjusted_matches:
            if course in student_prefs[student]:
                student_prefs[student][course] += weight
    
    flow = construct_flow(student_prefs, prof_prefs, course_info)

    status = flow.Solve()
    if status == flow.OPTIMAL:
        print('Solved flow with max value ', -flow.OptimalCost())

        courses = list([name for name, (target, _) in course_info.items() for _ in range(target)])
        write_matching(args.matchings, flow, students, courses, fixed_matches)
        
    else:
        print("Problem optimizing flow")