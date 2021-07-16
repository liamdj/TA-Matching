import csv
import argparse
from typing import Dict, List, Tuple
import numpy as np
import pandas as pd
from scipy import stats

import min_cost_flow

rank_scale = pd.Series(["Poor Matches", "Okay Matches", "Good Matches", "Excellent Matches"] , dtype="category")

ADVISOR_WEIGHT = 4
REPEAT_WEIGHT = 5
FAVORITE_WEIGHT = 0.1
FILL_COURSE_WEIGHT = 10
UNQUALIFIED_WEIGHT = -10

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
        # penalizes inflexible students
        scores *= len(scores) / len(courses)
        student_scores.append({c: s for c, s in zip(ranks.keys(), scores)})

    prof_scores = []
    for ranks in prof_prefs.values():
        scores = stats.zscore(list(ranks.values()))
        np.nan_to_num(scores, copy=False)
        # penalizes inflexible professors
        scores *= len(scores) / len(students)
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

def read_student_prefs_info(prefs_filename: str, info_filename: str, courses: List[str]) -> Dict[str, Tuple[Dict[str, int], List[str], str, str, float]]:
    ret = {}
    with open(prefs_filename, newline='') as file:
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
            ret[row["Name"]] = (rankings, repeated, advisors, favorite)

    with open(info_filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            repeated = row["Previous TA Experience"]
            advisors, _, _ = row["Advisor's Course"].partition(' (')

    return ret


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



def read_course_info(filename: str) -> Dict[str, Tuple[int, int, float]]:
    infos = {}
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            try:
                weight = FILL_COURSE_WEIGHT * float(row["Fill Weight"])
            except (ValueError, KeyError):
                weight = FILL_COURSE_WEIGHT
            infos[row["Course"]] = int(row["Target TAs"]), 0, weight

    return infos

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Generate matching from input data.")
    parser.add_argument("student_prefs", metavar="STUDENT PREFERENCES", help="csv file with student responses")
    parser.add_argument("prof_prefs", metavar="PROFESSOR PREFERENCES", help="csv file with professor responses")
    parser.add_argument("student_info", metavar="STUDENT INFORMATION", help="csv file with student information")
    parser.add_argument("course_info", metavar="COURSE INFORMATION", help="csv file with course information")
    parser.add_argument("--fixed", metavar="FIXED INPUT", help="csv file with required student-course matchings")
    parser.add_argument("--adjusted", metavar="ADJUSTED INPUT", help="csv file with adjustment weights for student-course matchings")
    parser.add_argument("--matchings", metavar="MATCHING OUTPUT", default='matchings.csv', help="location to write matching output")
    parser.add_argument("--additional", metavar="ADDITIONAL OUTPUT", nargs='?', const='add_TA.csv', help="location to write effects of gaining an additional TA for each course")
    parser.add_argument("--removal", metavar="REMOVAL OUTPUT", nargs='?', const='remove_TA.csv', help="location to write effects of removing each TA")
    args = parser.parse_args()

    course_info = read_course_info(args.course_info)
    student_prefs = read_student_prefs(args.student_prefs, course_info.keys())
    student_info = read_student_info(args.student_info)

    if args.fixed:
        fixed_matches = read_partial_matching(args.fixed)
        for student, course, _ in fixed_matches:
            student_prefs.pop(student, None)
            if course in course_info:
                cap, filled, weight = course_info[course]
                course_info[course] = cap, filled + 1, weight
    else:
        fixed_matches = []

    students = list(student_prefs.keys())
    prof_prefs = read_prof_prefs(args.prof, students, list(course_info.keys()))

    adjusted_matches = read_partial_matching(args.adjusted) if args.adjusted else []
    weights = weights_from_prefs(student_prefs, prof_prefs, adjusted_matches)

    flow = min_cost_flow.construct_flow(weights, course_info)

    status = flow.Solve()
    if status == flow.OPTIMAL:
        print('Solved flow with max value ', -flow.OptimalCost())

        course_slots = list([name for name, (cap, filled, _) in course_info.items() for _ in range(cap - filled)])
        write_matching(args.matchings, flow, student_prefs, prof_prefs, course_slots, fixed_matches)

    else:
        print("Problem optimizing flow")
        for arc in range(flow.NumArcs()):
            print("start: {}, end: {}, capacity: {}, cost: {}".format(flow.Tail(arc), flow.Head(arc), flow.Capacity(arc), flow.UnitCost(arc)))

    if args.additional:
        value_change = {}
        for c in course_info:
            # fill one additional course slot
            course_info_rm = course_info.copy()
            cap, filled, weight = course_info_rm[c]
            course_info_rm[c] = cap, filled + 1, weight

            flow_edit = min_cost_flow.construct_flow(weights, course_info_rm)
            value_change[c] = flow.OptimalCost() - flow_edit.OptimalCost() + min_cost_flow.fill_value(cap, filled, weight) if flow_edit.Solve() == flow_edit.OPTIMAL else -100
        
        with open(args.additional, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Course with Additional TA", "Change in Value"])
            for course, change in sorted(value_change.items(), key=lambda item: -item[1]):
                writer.writerow([course, change if change != -100 else "Error"])


    if args.removal:
        value_change = {}
        for i in range(len(students)):
            # remove student node
            weights_rm = np.delete(weights, i, 0)

            flow_edit = min_cost_flow.construct_flow(weights_rm, course_info)
            value_change[students[i]] = flow.OptimalCost() - flow_edit.OptimalCost() if flow_edit.Solve() == flow_edit.OPTIMAL else -100

        with open(args.removal, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Removed TA", "Change in Value"])
            for student, change in sorted(value_change.items(), key=lambda item: -item[1]):
                writer.writerow([student, change if change != -100 else "Error"])
