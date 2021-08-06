import csv
import sys
import argparse
from typing import Tuple

from pandas.core.series import Series
import numpy as np
import pandas as pd
from scipy import stats

from min_cost_flow import MatchingGraph

rank_scale = ["Unqualified", "Poor", "Okay", "Good", "Excellent"]

ADVISOR_WEIGHT = 3
REPEAT_WEIGHT = 3
FAVORITE_WEIGHT = 0.2
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

    weights = np.full((len(student_info.index), len(course_info.index)), np.nan)
    for si, student in enumerate(student_info.index):
        s_scores = student_scores.loc[student]
        previous = student_info.loc[student, "Previous Courses"].split(';')
        advisor = student_info.loc[student, "Advisor's Course"]
        favorite = student_info.loc[student, "Favorite Course"]
        for ci, course in enumerate(course_info.index):
            c_scores = course_scores.loc[course]
            if not pd.isna(c_scores[student]) and not pd.isna(s_scores[course]):
                weights[si, ci] = s_scores[course] + c_scores[student]

                if course in previous:
                    weights[si, ci] += REPEAT_WEIGHT
                if course == advisor:
                    weights[si, ci] += ADVISOR_WEIGHT
                if course == favorite:
                    weights[si, ci] += FAVORITE_WEIGHT

    return weights

def read_student_responses(filename: str) -> Tuple[pd.DataFrame, pd.Series]:
    # expand list of courses into ranking for each course
    rank_rows = []
    names = []
    favorites = []
    netids = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {}
            for rank in rank_scale:
                for course in row[rank + " Matches"].split(';'):
                    if course:
                        collect[course] = rank
            rank_rows.append(pd.Series(collect, name=row["Name"]))
            names.append(row["Name"])
            favorites.append(row["Favorite Course"])
            id, _, _ = row["Username"].partition('@')
            netids.append(id)

    return pd.concat(rank_rows, axis='columns').fillna("None").T, pd.Series(favorites, index=names), pd.Series(netids, index=names)

def read_prof_responses(filename: str) -> pd.DataFrame:
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

def check_student_input(info: pd.Series, responses: pd.Series):

    for index in info[info.index.duplicated()].index:
        sys.exit("Duplicate information for {}. Exiting without solving".format(index))
    for index in responses[responses.index.duplicated()].index:
        sys.exit("Duplicate responses for {}. Exiting without solving".format(index))
    
    df = pd.concat([info, responses], axis=1)
    for index, row in df.iloc[:, 0].compare(df.iloc[:, 1]).iterrows():
        if pd.isna(row["self"]):
            sys.exit("Responses for {} are given, but the accompanying information is missing. Exiting without solving.".format(index))
        elif pd.isna(row["other"]):
            sys.exit("Responses for {} are missing. Exiting without solving.".format(index))
        else:
            print("'{}' is given as the net id given for {}, but the respondent used '{}'.".format(row["self"], index, row["other"]))


    
def check_course_input(info: pd.Index, responses: pd.Index):

    for index in info[info.duplicated()]:
        sys.exit("Duplicate information for {}. Exiting without solving.".format(index))
    for index in responses[responses.duplicated()]:
        sys.exit("Duplicate responses for {}. Exiting without solving.".format(index))

    for index in responses.difference(info):
        sys.exit("Responses for {} are given, but the accompanying information is missing. Exiting without solving.".format(index))
    for index in info.difference(responses):
        sys.exit("Responses for {} are missing. Exiting without solving.".format(index))


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Generate matching from input data.")
    parser.add_argument("--path", metavar="FOLDER", default='', help="prefix to file paths")
    parser.add_argument("--student_prefs", metavar="STUDENT PREFERENCES", default='student_prefs.csv', help="csv file with student responses")
    parser.add_argument("--prof_prefs", metavar="PROFESSOR PREFERENCES", default='prof_prefs.csv', help="csv file with professor responses")
    parser.add_argument("--student_info", metavar="STUDENT INFORMATION", default='student_info.csv', help="csv file with student information")
    parser.add_argument("--course_info", metavar="COURSE INFORMATION", default='course_info.csv',help="csv file with course information")
    parser.add_argument("--fixed", metavar="FIXED INPUT", help="csv file with required student-course matchings")
    parser.add_argument("--adjusted", metavar="ADJUSTED INPUT", help="csv file with adjustment weights for student-course matchings")
    parser.add_argument("--matchings", metavar="MATCHING OUTPUT", default='matchings.csv', help="location to write matching output")
    parser.add_argument("--additional", metavar="ADDITIONAL OUTPUT", nargs='?', const='add_TA.csv', help="location to write effects of gaining an additional TA for each course")
    parser.add_argument("--removal", metavar="REMOVAL OUTPUT", nargs='?', const='remove_TA.csv', help="location to write effects of removing each TA")
    args = parser.parse_args()


    student_info = pd.read_csv(args.path + args.student_info, index_col="Student", keep_default_na=False, dtype=str)
    student_info["Assign Weight"] = pd.to_numeric(student_info["Assign Weight"], errors='coerce', downcast='float').fillna(DEFAULT_ASSIGN_WEIGHT)

    course_info = pd.read_csv(args.path + args.course_info, index_col="Course", dtype = {"Course": str, "TA Slots": int, "Fill Weight": float})
    course_info.index = course_info.index.astype(str)
    course_info["TA Slots"].fillna(1, inplace=True)
    course_info["Fill Weight"].fillna(DEFAULT_FILL_WEIGHT, inplace=True)

    student_ranks, favorites, netids = read_student_responses(args.path + args.student_prefs)
    check_student_input(student_info["Net ID"], netids)
    student_info["Favorite Course"] = favorites

    course_ranks = read_prof_responses(args.path + args.prof_prefs)
    check_course_input(course_info.index, course_ranks.index)

    student_scores = generate_scores(student_ranks, course_info.index)
    course_scores = generate_scores(course_ranks, student_info.index)

    weights = match_weights(student_info, student_scores, course_info, course_scores)

    if args.adjusted:
        adjusted_matches = read_partial_matching(args.path + args.adjusted)
        for _, row in adjusted_matches.iterrows():
            si = student_info.index.get_loc(row["Student"])
            ci = course_info.index.get_loc(row["Course"])
            weights[si, ci] += row["Match Weight"]

    if args.fixed:
        fixed_matches = pd.read_csv(args.path + args.fixed, dtype=str)
        fixed_matches["Student index"] = [student_info.index.get_loc(student) for student in fixed_matches["Student"]]
        fixed_matches["Course index"] = [course_info.index.get_loc(course) for course in fixed_matches["Course"]]
    else:
        fixed_matches = pd.DataFrame(columns=["Student", "Course", "Student index", "Course index"])

    graph = MatchingGraph(weights, student_info["Assign Weight"], course_info["Fill Weight"], course_info["TA Slots"], fixed_matches)

    if graph.solve():
        print('Successfully solved flow')

        graph.write_matching(args.path + args.matchings, weights, student_info, student_ranks, student_scores, course_info.index, course_ranks, course_scores, fixed_matches)

    else:
        print("Problem optimizing flow")
        # print(graph.flow.BAD_COST_RANGE, graph.flow.BAD_RESULT, graph.flow.INFEASIBLE, graph.flow.NOT_SOLVED, graph.flow.UNBALANCED)
        graph.print()

    if args.additional:
        value_change = {}
        for ci, course in enumerate(course_info.index):
            if ci not in fixed_matches["Course index"]:
                # fill one additional course slot
                fixed_edit = pd.concat([fixed_matches.T, pd.Series(["", course, -1, ci], index=["Student", "Course", "Student index", "Course index"])], axis=1).T

                graph_edit = MatchingGraph(weights, student_info["Assign Weight"], course_info["Fill Weight"], course_info["TA Slots"], fixed_edit)
                value_change[course] = graph.flow.OptimalCost() - graph_edit.flow.OptimalCost() if graph_edit.solve() else -100
        
        with open(args.path + args.additional, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Course with Additional TA", "Change in Value"])
            for course, change in sorted(value_change.items(), key=lambda item: -item[1]):
                writer.writerow([course, change if change != -100 else "Error"])


    if args.removal:
        value_change = {}
        for si, student in enumerate(student_info.index):
            if si not in fixed_matches["Student index"]:
                # no edges from student node
                fixed_edit = pd.concat([fixed_matches.T, pd.Series([student, "", si, -1], index=["Student", "Course", "Student index", "Course index"])], axis=1).T

                graph_edit = MatchingGraph(weights, student_info["Assign Weight"], course_info["Fill Weight"], course_info["TA Slots"], fixed_edit)
                value_change[student] = graph.flow.OptimalCost() - graph_edit.flow.OptimalCost() if graph_edit.solve() else -100

        with open(args.path + args.removal, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Removed TA", "Change in Value"])
            for student, change in sorted(value_change.items(), key=lambda item: -item[1]):
                writer.writerow([student, change if change != -100 else "Error"])
