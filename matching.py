import csv
import argparse
from typing import Dict, List, Tuple
import numpy as np
from scipy import stats
from ortools.graph import pywrapgraph

DIGITS = 3

rank_scale = ["Poor Matches", "Okay Matches", "Good Matches", "Excellent Matches"] 

ADVISOR_WEIGHT = 4
REPEAT_WEIGHT = 5
FAVORITE_WEIGHT = 0.1
FILL_COURSE_WEIGHT = 10
UNQUALIFIED_WEIGHT = -10


def construct_flow(weights, course_info):

    N, M = weights.shape[0], sum([target for target, _ in course_info.values()])
    source, sink = N + M, N + M + 1
    flow = pywrapgraph.SimpleMinCostFlow(2 + N + M)

    # Each student cannot fill >1 course slot
    for i in range(N):
        flow.AddArcWithCapacityAndUnitCost(source, i, 1, 0)

    # Each course slot cannot have >1 TA
    # Value of filling slot is reciprocal with slot index
    node = N
    for target, value in course_info.values():
        for i in range(1, target + 1):
            cost = int(-value / i * 10 ** DIGITS)
            flow.AddArcWithCapacityAndUnitCost(node, sink, 1, cost)
            node += 1

    # Edge weights given by preferences
    for si in range(N):
        node = N
        for ci, info in enumerate(course_info.values()):
            target, _ = info
            if weights[si, ci] > -5:
                cost = int(-weights[si, ci] * 10 ** DIGITS)
                for slot in range(target):
                    flow.AddArcWithCapacityAndUnitCost(si, node + slot, 1, cost)
            node += target

    # Max number of slots must be filled
    flow.SetNodeSupply(source, min(N, M))
    flow.SetNodeSupply(sink, -min(N, M))

    return flow

def write_matching(filename: str, flow, student_prefs, prof_prefs, courses, fixed_matches):
    students = list(student_prefs.keys())
    with open(filename, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Student", "Course", "Student Preference", "Professor Preference"])

        for arc in range(flow.NumArcs()):
            si, ci = flow.Tail(arc), flow.Head(arc) - len(students)
            if flow.Flow(arc) > 0 and si < len(students):
                student = students[si]
                course = courses[ci]
                s_rank, _, _ = rank_scale[student_prefs[student][0][course]].partition(' ')
                c_rank, _, _ = rank_scale[prof_prefs[course][student]].partition(' ')
                writer.writerow([student, course, s_rank, c_rank])

        for student, course, _ in fixed_matches:
            writer.writerow([student, course, "Fixed", ""])

def read_partial_matching(filename: str) -> List[Tuple[str, str, float]]:
    matchings = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            weight = float(row["Weight"]) if "Weight" in row else 0
            matchings.append((row["Student"], row["Course"], weight))

    return matchings

def weights_from_prefs(student_prefs: dict, prof_prefs: dict, adjusted_weights: list=[]) -> np.array:

    students = list(student_prefs.keys())
    courses = list(prof_prefs.keys())

    student_scores = []
    for ranks, _, _, _ in student_prefs.values():
        scores = stats.zscore(list(ranks.values()))
        np.nan_to_num(scores, copy=False)
        student_scores.append({c: s for c, s in zip(ranks.keys(), scores)})

    prof_scores = []
    for ranks in prof_prefs.values():
        scores = stats.zscore(list(ranks.values()))
        np.nan_to_num(scores, copy=False)
        prof_scores.append({c: s for c, s in zip(ranks.keys(), scores)})

    weights = UNQUALIFIED_WEIGHT * np.ones((len(student_prefs), len(prof_prefs)))
    for si, (student, s_scores) in enumerate(zip(students, student_scores)):
        _, previous, advisor, favorite = student_prefs[student]
        for ci, (course, c_scores) in enumerate(zip(courses, prof_scores)):
            if student in c_scores and course in s_scores:
                weights[si, ci] = s_scores[course] + c_scores[student]
            if course in previous:
                weights[si, ci] += REPEAT_WEIGHT
            if course == advisor:
                weights[si, ci] += ADVISOR_WEIGHT
            if course == favorite:
                weights[si, ci] += FAVORITE_WEIGHT
 
    for student, course, weight in adjusted_weights:
        si = students.index(student)
        ci = courses.index(course)
        weights[si, ci] += weight

    return weights

def read_student_prefs(filename: str, courses: List[str]) -> Dict[str, Tuple[Dict[str, int], List[str], str, str]]:
    prefs = {}
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            rankings = dict.fromkeys(courses, 1)
            for value, question in enumerate(rank_scale):
                for course in row[question].split(';'):
                    if course:
                        rankings[course] = value
            for course in row["Unqualified Courses"].split(';'):
                if course:
                    del rankings[course]

            # netid, _, _ = row["Username"].partition('@')
            repeated = row["Previous TA Experience"]
            advisors, _, _ = row["Advisor's Course"].partition(' (')
            favorite = row["Favorite Course"]
            prefs[row["Name"]] = (rankings, repeated, advisors, favorite)

    return prefs


def read_prof_prefs(filename: str, students: List[str], courses: List[str]) -> Dict[str, Dict[str, float]]:
    prefs = {c: dict.fromkeys(students, 1) for c in courses}

    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            rankings = prefs[row["Course"]]
            for value, question in enumerate(rank_scale):
                for student in row[question].split(';'):
                    if student:
                        rankings[student] = value
            for student in row["Unqualified Students"].split(';'):
                if student:
                    del rankings[student]
    
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

    course_info = read_course_info(args.course)
    student_prefs = read_student_prefs(args.student, course_info.keys())

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

    weights = weights_from_prefs(student_prefs, prof_prefs)

    flow = construct_flow(weights, course_info)

    status = flow.Solve()
    print(weights)
    print(student_prefs)
    print(prof_prefs)
    if status == flow.OPTIMAL:
        print('Solved flow with max value ', -flow.OptimalCost())

        course_slots = list([name for name, (target, _) in course_info.items() for _ in range(target)])
        write_matching(args.matchings, flow, student_prefs, prof_prefs, course_slots, fixed_matches)

    else:
        print("Problem optimizing flow")
        for arc in range(flow.NumArcs()):
            print("start: {}, end: {}, capacity: {}, cost: {}".format(flow.Tail(arc), flow.Head(arc), flow.Capacity(arc), flow.UnitCost(arc)))


    if args.additional:
        for c in course_info:
            pass
