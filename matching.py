import csv
import argparse
from typing import Dict, List, Tuple
import numpy as np
import pandas as pd
from scipy import stats

from min_cost_flow import MatchingGraph

rank_scale = ["Unqualified", "Poor", "Okay", "Good", "Excellent"]

ADVISOR_WEIGHT = 4
REPEAT_WEIGHT = 5
FAVORITE_WEIGHT = 0.1
DEFAULT_FILL_WEIGHT = 10
DEFAULT_ASSIGN_WEIGHT = 0
BASE_MATCH_WEIGHT = 10

def read_partial_matching(filename: str) -> pd.DataFrame:
    return pd.read_csv(filename) if filename else pd.DataFrame()

# generator that skips values of unqualified
def qual_resp(series, cols):
    for c in cols:
        v = 2 if series[c] == "None" else rank_scale.index(series[c])
        if v > 0:
            yield c, v

def generate_scores(data: pd.DataFrame, cols: pd.Series) -> pd.DataFrame:
    rows = []
    for index, row in data.iterrows():
        # collect values for qualified responses
        values = [v for _, v in qual_resp(row, cols)]
        # normalized scores associated with rankings
        scores = np.nan_to_num(stats.zscore(values))
        # penalize inflexible students / profs
        scores *= len(values) / len(cols)
        rows.append(pd.Series({c: score for (c, _), score in zip(qual_resp(row, cols), scores)}, name=index))
    return pd.concat(rows, axis="columns").T

def match_weights(student_info: pd.DataFrame, student_scores: pd.DataFrame, course_info: pd.DataFrame, course_scores: pd.DataFrame) -> np.array:

    weights = np.zeros((len(student_info.index), len(course_info.index)))
    for si, student in enumerate(student_info.index):
        s_scores = student_scores.loc[student]
        previous = student_info.loc[student, "Previous Courses"].split(';')
        advisor = student_info.loc[student, "Advisor's Course"]
        favorite = student_info.loc[student, "Favorite Course"]
        for ci, course in enumerate(course_info.index):
            c_scores = course_scores.loc[course]
            if not pd.isna(c_scores[student]) and not pd.isna(s_scores[course]):
                weights[si, ci] = s_scores[course] + c_scores[student] + BASE_MATCH_WEIGHT

                if course in previous:
                    weights[si, ci] += REPEAT_WEIGHT
                if course == advisor:
                    weights[si, ci] += ADVISOR_WEIGHT
                if course == favorite:
                    weights[si, ci] += FAVORITE_WEIGHT

    return weights

def read_student_responses(filename: str, courses: List[str]) -> Tuple[pd.DataFrame, pd.Series]:
    # expand list of courses into ranking for each course
    rank_rows = []
    favorites = {}
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {}
            for rank in rank_scale:
                for course in row[rank + " Matches"].split(';'):
                    if course:
                        collect[course] = rank
            # netid, _, _ = row["Username"].partition('@') (')
            rank_rows.append(pd.Series(collect, name=row["Name"]))
            favorites[row["Name"]] = row["Favorite Course"]

    return pd.concat(rank_rows, axis='columns').fillna("None").T, pd.Series(favorites)
    return pd.concat([responses, info], axis='columns')

def read_prof_responses(filename: str, students: List[str]) -> pd.DataFrame:
    # expand list of students into ranking for each student
    rank_rows = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {}
            for rank in rank_scale:
                for student in row[rank + " Matches"].split(';'):
                    if student:
                        collect[student] = rank
            rank_rows.append(rank_rows.append(pd.Series(collect, name=row["Course"])))

    return pd.concat(rank_rows, axis='columns').fillna("None").T
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

    student_ranks, favorites = read_student_responses(args.student_prefs, list(course_info.index))
    student_info["Favorite Course"] = favorites

    course_ranks = read_prof_responses(args.prof_prefs, list(student_info.index))

    student_scores = generate_scores(student_ranks, course_info.index)
    course_scores = generate_scores(course_ranks, student_info.index)

    weights = match_weights(student_info, student_scores, course_info, course_scores)

    adjusted_matches = read_partial_matching(args.adjusted)
    for _, row in adjusted_matches.iterrows():
        si = student_info.index.get_loc(row["Student"])
        ci = course_info.index.get_loc(row["Course"])
        weights[si, ci] += row["Match Weight"]

    fixed_matches = read_partial_matching(args.fixed)
    fixed_matches["Student index"] = [student_info.index.get_loc(row["Student"]) for _, row in fixed_matches.iterrows()]
    fixed_matches["Course index"] = [course_info.index.get_loc(row["Course"]) for _, row in fixed_matches.iterrows()]

    graph = MatchingGraph(weights, student_info["Assign Weight"], course_info["Fill Weight"], course_info["TA Slots"], fixed_matches)

    if graph.solve():
        print('Successfully solved flow')
        graph.write(args.matchings, student_info.index, student_ranks, student_scores, course_info.index, course_ranks, course_scores, fixed_matches)

    else:
        print("Problem optimizing flow")
        # print(graph.flow.BAD_COST_RANGE, graph.flow.BAD_RESULT, graph.flow.INFEASIBLE, graph.flow.NOT_SOLVED, graph.flow.UNBALANCED)
        graph.print()

    # if args.additional:
    #     value_change = {}
    #     for c in course_info:
    #         # fill one additional course slot
    #         course_info_rm = course_info.copy()
    #         cap, filled, weight = course_info_rm[c]
    #         course_info_rm[c] = cap, filled + 1, weight

    #         flow_edit = min_cost_flow.construct_flow(weights, course_info_rm)
    #         value_change[c] = flow.OptimalCost() - flow_edit.OptimalCost() + min_cost_flow.fill_value(cap, filled, weight) if flow_edit.Solve() == flow_edit.OPTIMAL else -100
        
    #     with open(args.additional, 'w', newline='') as file:
    #         writer = csv.writer(file)
    #         writer.writerow(["Course with Additional TA", "Change in Value"])
    #         for course, change in sorted(value_change.items(), key=lambda item: -item[1]):
    #             writer.writerow([course, change if change != -100 else "Error"])


    # if args.removal:
    #     value_change = {}
    #     for i in range(len(students)):
    #         # remove student node
    #         weights_rm = np.delete(weights, i, 0)

    #         flow_edit = min_cost_flow.construct_flow(weights_rm, course_info)
    #         value_change[students[i]] = flow.OptimalCost() - flow_edit.OptimalCost() if flow_edit.Solve() == flow_edit.OPTIMAL else -100

    #     with open(args.removal, 'w', newline='') as file:
    #         writer = csv.writer(file)
    #         writer.writerow(["Removed TA", "Change in Value"])
    #         for student, change in sorted(value_change.items(), key=lambda item: -item[1]):
    #             writer.writerow([student, change if change != -100 else "Error"])
