import csv
import argparse
from typing import Dict, List, Tuple
import numpy as np
import pandas as pd
from scipy import stats

import min_cost_flow

rank_scale = ["Unqualified", "Poor", "Okay", "Good", "Excellent"]

ADVISOR_WEIGHT = 4
REPEAT_WEIGHT = 5
FAVORITE_WEIGHT = 0.1
DEFAULT_FILL_WEIGHT = 10
DEFAULT_ASSIGN_WEIGHT = 0
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

def read_partial_matching(filename: str) -> pd.DataFrame:
    return pd.read_csv(filename) if filename else pd.DataFrame()

# generator that skips values of unqualified
def qual_resp(series, cols):
    for c in cols:
        if series[c] > 0:
            yield c, series[c]

def weights_from_data(student_data: pd.DataFrame, course_data: pd.DataFrame, adjusted_weights: pd.DataFrame) -> np.array:

    student_scores = []
    for student, row in student_data.iterrows():
        # collect values for qualified responses
        values = [v for _, v in qual_resp(row, course_data.index)]
        # normalize scores associated with rankings
        scores = np.nan_to_num(stats.zscore(values))
        # penalize inflexible students
        scores *= len(values) / len(course_data.index)
        student_scores.append({course: score for (course, _), score in zip(qual_resp(row, course_data.index), scores)})


    course_scores = []
    for course, row in course_data.iterrows():
        # collect values for qualified responses
        values = [v for _, v in qual_resp(row, student_data.index)]
        # normalize scores associated with rankings
        scores = np.nan_to_num(stats.zscore(values))
        # penalize inflexible profs
        scores *= len(values) / len(student_data.index)
        course_scores.append({student: score for (student, _), score in zip(qual_resp(row, student_data.index), scores)})

    weights = UNQUALIFIED_WEIGHT * np.ones((len(student_data.index), len(course_data.index)))
    for si, ((student, row), s_scores) in enumerate(zip(student_data.iterrows(), student_scores)):
        previous = row["Previous Courses"].split(';')
        for ci, (course, c_scores) in enumerate(zip(course_data.index, course_scores)):
            if student in c_scores and course in s_scores:
                weights[si, ci] = s_scores[course] + c_scores[student]
            if course in previous:
                weights[si, ci] += REPEAT_WEIGHT
            if course == row["Advisor's Course"]:
                weights[si, ci] += ADVISOR_WEIGHT
            if course == row["Favorite Course"]:
                weights[si, ci] += FAVORITE_WEIGHT
 
    for _, row in adjusted_weights.iterrows():
        si = student_data.index.get_loc(row["Student"])
        ci = course_data.index.get_loc(row["Course"])
        weights[si, ci] += row["Match Weight"]

    return weights

def read_student_data(response_filename: str, info: pd.DataFrame, courses: List[str]) -> pd.DataFrame:
    # expand list of courses into ranking for each course
    ls = []
    with open(response_filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            series = pd.Series(2, index=courses, name=row["Name"])
            for value, question in enumerate(rank_scale):
                for course in row[question + " Matches"].split(';'):
                    if course:
                        series[course] = value
            # netid, _, _ = row["Username"].partition('@') (')
            series["Favorite Course"] = row["Favorite Course"]
            ls.append(series)

    responses = pd.concat(ls, axis='columns').T
    return pd.concat([responses, info], axis='columns')

def read_course_data(response_filename: str, info: pd.DataFrame, students: List[str]) -> pd.DataFrame:
    # expand list of students into ranking for each student
    ls = []
    with open(response_filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            series = pd.Series(2, index=students, name=row["Course"])
            for value, question in enumerate(rank_scale):
                for student in row[question + " Matches"].split(';'):
                    if student:
                        series[student] = value
            ls.append(series)

    responses = pd.concat(ls, axis='columns').T
    return pd.concat([responses, info], axis='columns')

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


    student_info = pd.read_csv(args.student_info, index_col="Student", keep_default_na=False, dtype=str)
    student_info["Assign Weight"] = pd.to_numeric(student_info["Assign Weight"], errors='coerce', downcast='float').fillna(DEFAULT_ASSIGN_WEIGHT)

    course_info = pd.read_csv(args.course_info, index_col="Course", dtype = {"Course": str, "TA Slots": int, "Fill Weight": float})
    course_info["TA Slots"].fillna(1, inplace=True)
    course_info["Fill Weight"].fillna(DEFAULT_FILL_WEIGHT, inplace=True)

    student_data = read_student_data(args.student_prefs, student_info, list(course_info.index))
    course_data = read_course_data(args.prof_prefs, course_info, list(student_info.index))

    adjusted_matches = read_partial_matching(args.adjusted)
    weights = weights_from_data(student_data, course_data, adjusted_matches)

    fixed_matches = read_partial_matching(args.fixed)
    flow = min_cost_flow.construct_flow(weights, course_info)

    status = flow.Solve()
    if status == flow.OPTIMAL:
        print('Solved flow with max value ', -flow.OptimalCost())

        course_slots = list([name for name, (cap, filled, _) in course_info.items() for _ in range(cap - filled)])
        write_matching(args.matchings, flow, student_info, course_info, course_slots, fixed_matches)

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
