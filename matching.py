import csv
import sys
import os
import argparse
from typing import Tuple

import numpy as np
import pandas as pd
from scipy import stats

import min_cost_flow

student_data_cols = ['Name', 'Weight', 'Previous', 'Advisors']
student_options = ['Okay', 'Good', 'Favorite']
course_data_cols = ['Slots', 'Weight']
course_options = ['Veto', 'Favorite']

PREVIOUS_WEIGHT = 2
ADVISORS_WEIGHT = 2
STUDENT_PREF_WEIGHT = 1
PROF_PREF_WEIGHT = 1.5

DEFAULT_FILL_WEIGHT = 10
DEFAULT_ASSIGN_WEIGHT = 10


# generator for values for courses that student indicated is qualified for
def student_rankings(series, cols):
    for c in cols:
        if c in series and pd.notna(series[c]):
            yield c, student_options.index(series[c])

def get_student_scores(data: pd.DataFrame, cols: pd.Index) -> pd.DataFrame:

    rows = []
    for index, row in data.iterrows():
        # collect values for qualified responses
        values = [v for _, v in student_rankings(row, cols)]
        # normalized scores associated with rankings
        scores = np.nan_to_num(stats.zscore(values))
        rows.append(pd.Series({c: score for (c, _), score in zip(student_rankings(row, cols), scores)}, name=index))
    return pd.concat(rows, axis='columns').T

# generator for values for students that prof indicated is qualified for
def course_rankings(series, cols):
    for c in cols:
        if c not in series or pd.isna(series[c]):
            yield c, 0
        else: 
            v = course_options.index(series[c])
            if v > 0:
                yield c, v

def get_course_scores(data: pd.DataFrame, cols: pd.Index) -> pd.DataFrame:

    rows = []
    for index, row in data.iterrows():
        # collect values for qualified responses
        values = [v for _, v in course_rankings(row, cols)]
        # normalized scores associated with rankings
        scores = np.nan_to_num(stats.zscore(values))
        rows.append(pd.Series({c: score for (c, _), score in zip(course_rankings(row, cols), scores)}, name=index))
    return pd.concat(rows, axis='columns').T

def match_weights(student_data: pd.DataFrame, student_scores: pd.DataFrame, course_scores: pd.DataFrame) -> np.array:

    weights = np.full((len(student_scores.index), len(course_scores.index)), np.nan)
    for si, (student, s_scores) in enumerate(student_scores.iterrows()):
        previous = student_data.loc[student, 'Previous'].split(';')
        advisor = student_data.loc[student, 'Advisors'].split(';')
        for ci, (course, c_scores) in enumerate(course_scores.iterrows()):
            if not pd.isna(c_scores[student]) and not pd.isna(s_scores[course]):
                weights[si, ci] = s_scores[course] * STUDENT_PREF_WEIGHT + c_scores[student] * PROF_PREF_WEIGHT
                if course in previous:
                    weights[si, ci] += PREVIOUS_WEIGHT
                if course in advisor:
                    weights[si, ci] += ADVISORS_WEIGHT
    return weights

def read_student_data(filename: str) -> pd.DataFrame:
    rank_rows = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {col: row[col] for col in student_data_cols}
            # expand list of courses into ranking for each course
            for rank in student_options:
                for course in row[rank].split(';'):
                    if course:
                        collect[course] = rank
            rank_rows.append(pd.Series(collect, name=row['Netid']))

    df = pd.concat(rank_rows, axis='columns').T
    df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce', downcast='float').fillna(DEFAULT_ASSIGN_WEIGHT)
    return df

def read_course_data(filename: str) -> pd.DataFrame:
    rank_rows = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {col : row[col] for col in course_data_cols}
             # expand list of students into ranking for each student
            for rank in course_options:
                for student in row[rank].split(';'):
                    if student:
                        collect[student] = rank
            rank_rows.append(rank_rows.append(pd.Series(collect, name=row['Course'])))

    df = pd.concat(rank_rows, axis='columns').T
    df['Slots'] = pd.to_numeric(df['Slots'], errors='coerce', downcast='integer').fillna(1)
    df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce', downcast='float').fillna(DEFAULT_ASSIGN_WEIGHT)
    return df

def check_input(student_data: pd.Series, course_data: pd.Series):
    for index in student_data[student_data.index.duplicated()].index:
        sys.exit('Duplicate rows for netid {}. Exiting without solving'.format(index))
    for index in course_data[course_data.index.duplicated()].index:
        sys.exit('Duplicate rows for course {}. Exiting without solving'.format(index))

def test_additional_TA(path, student_data, course_data, weights, fixed_matches):
    value_change = {}
    for ci, course in enumerate(course_data.index):
        # fill one additional course slot
        fixed_edit = pd.concat([fixed_matches.T, pd.Series(['', course, -1, ci], index=['Student', 'Course', 'Student index', 'Course index'])], axis=1).T
        graph_edit = min_cost_flow.MatchingGraph(weights, student_data['Weight'], course_data['Weight'], course_data['Slots'], fixed_edit)
        value_change[course] = (graph.flow.OptimalCost() - graph_edit.flow.OptimalCost()) / 10 ** min_cost_flow.DIGITS if graph_edit.solve() else -100
    
    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Course with Additional TA', 'Change in Value'])
        for course, change in sorted(value_change.items(), key=lambda item: -item[1]):
            writer.writerow([course, change if change != -100 else 'Error'])

def test_removing_TA(path, student_data, course_data, weights, fixed_matches):
    value_change = {}
    for si, student in enumerate(student_data.index):
        if si not in fixed_matches['Student index']:
            # no edges from student node
            fixed_edit = pd.concat([fixed_matches.T, pd.Series([student, '', si, -1], index=['Student', 'Course', 'Student index', 'Course index'])], axis=1).T
            graph_edit = min_cost_flow.MatchingGraph(weights, student_data['Weight'], course_data['Weight'], course_data['Slots'], fixed_edit)
            value_change[student] = (graph.flow.OptimalCost() - graph_edit.flow.OptimalCost()) / 10 ** min_cost_flow.DIGITS if graph_edit.solve() else -100

    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Removed TA', 'Change in Value'])
        for student, change in sorted(value_change.items(), key=lambda item: -item[1]):
            writer.writerow([student, change if change != -100 else 'Error'])

def validate_path_args(args):
    # ensure trailing slash on path directory
    if args.path[-1] != '/':
        args.path += '/'
        
    # ensure output directory exists
    output_path = args.path + args.output
    os.makedirs(output_path, exist_ok=True)


if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Generate matching from input data.')
    parser.add_argument('--path', metavar='FOLDER', default='', help='prefix to file paths')
    parser.add_argument('--student_data', metavar='STUDENT DATA', default='inputs/student_data.csv', help='csv file with student data rows')
    parser.add_argument('--course_data', metavar='COURSE DATA', default='inputs/course_data.csv', help='csv file with course data rows')
    parser.add_argument('--fixed', metavar='FIXED INPUT', default='inputs/fixed.csv', help='csv file with fixed student-course matchings')
    parser.add_argument('--adjusted', metavar='ADJUSTED INPUT', default='inputs/adjusted.csv', help='csv file with adjustment weights for student-course matchings')
    parser.add_argument('--output', metavar='MATCHING OUTPUT', default='outputs/', help='location to write matching output')
    args = parser.parse_args()
    validate_path_args(args)

    student_data = read_student_data(args.path + args.student_data)
    course_data = read_course_data(args.path + args.course_data)

    for student in student_data.index:
        if student not in course_data.columns.values:
            course_data[student] = np.nan
    for course in course_data.index:
        if course not in student_data.columns.values:
            student_data[course] = np.nan   

    student_scores = get_student_scores(student_data, course_data.index)
    course_scores = get_course_scores(course_data, student_data.index)

    weights = match_weights(student_data, student_scores, course_scores)

    if os.path.isfile(args.path + args.adjusted):
        adjusted_matches = pd.read_csv(args.path + args.adjusted, dtype={'Netid': str, 'Course': str, 'Weight': float})
        for _, row in adjusted_matches.iterrows():
            si = student_scores.index.get_loc(row['Netid'])
            ci = course_scores.index.get_loc(row['Course'])
            weights[si, ci] += row['Weight']

    if os.path.isfile(args.path + args.fixed):
        fixed_matches = pd.read_csv(args.path + args.fixed, dtype=str)
        fixed_matches['Student index'] = [student_scores.index.get_loc(student) for student in fixed_matches['Netid']]
        fixed_matches['Course index'] = [course_scores.index.get_loc(course) for course in fixed_matches['Course']]
    else:
        fixed_matches = pd.DataFrame(columns=['Netid', 'Course', 'Student index', 'Course index'])

    graph = min_cost_flow.MatchingGraph(weights, student_data['Weight'], course_data['Weight'], course_data['Slots'], fixed_matches)

    if graph.solve():
        print('Successfully solved flow')

        graph.write_matching(args.path + args.output + 'matchings.csv', weights, student_data, student_scores, course_data, course_scores, fixed_matches)

    else:
        print('Problem optimizing flow')
        graph.print()

    test_additional_TA(args.path + args.output + 'additional_TA.csv', student_data, course_data, weights, fixed_matches)
    test_removing_TA(args.path + args.output + 'remove_TA.csv', student_data, course_data, weights, fixed_matches)
