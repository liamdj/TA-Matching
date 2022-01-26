from collections import defaultdict
import csv
import sys
import os
import argparse
from typing import List, Tuple

import numpy as np
import pandas as pd
from scipy import stats
import params

import min_cost_flow

student_data_cols = ['Name', 'Weight', 'Year', 'Bank', 'Join', 'Previous',
                     'Advisors', 'Notes']
student_options = ['Okay', 'Good', 'Favorite']
course_data_cols = ['Slots', 'Weight', 'Instructor']
course_options = ['Veto', 'Favorite']


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
        # AF and MW changed this next line:
        # scores = np.nan_to_num(stats.zscore(values))
        scores = [float(v) for v in values]
        rows.append(
            pd.Series(
                {c: score for (c, _), score in
                 zip(student_rankings(row, cols), scores)}, name=index,
                dtype=np.float64))
    return pd.concat(rows, axis='columns').T


def course_rankings(series, cols):
    """generator for values for students that prof indicated is qualified for"""
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
        # AF and MW changed this next line:
        # scores = np.nan_to_num(stats.zscore(values))
        scores = [float(v) for v in values]
        rows.append(
            pd.Series(
                {c: score for (c, _), score in
                 zip(course_rankings(row, cols), scores)}, name=index))
    return pd.concat(rows, axis='columns').T


def match_weights(student_data: pd.DataFrame, student_scores: pd.DataFrame,
                  course_data: pd.DataFrame,
                  course_scores: pd.DataFrame) -> np.array:
    weights = np.full(
        (len(student_scores.index), len(course_scores.index)), np.nan)
    for si, (student, s_scores) in enumerate(student_scores.iterrows()):
        previous = student_data.loc[student, 'Previous'].split(';')
        advisors = student_data.loc[student, 'Advisors'].split(';')
        for ci, (course, c_scores) in enumerate(course_scores.iterrows()):
            instructors = course_data.loc[course, 'Instructor'].split(';')
            if not pd.isna(
                    c_scores[student]) and course in s_scores and not pd.isna(
                s_scores[course]):
                weights[si, ci] = s_scores[course] * params.STUDENT_PREF + \
                                  c_scores[student] * params.PROF_PREF
                if student_data.loc[student, course] == 'Favorite' and c_scores[
                    student] > 0:
                    weights[si, ci] += params.FAVORITE
                if student_data.loc[student, course] == 'Okay':
                    weights[si, ci] += params.OKAY_COURSE_PENALTY
                if course in previous:
                    weights[si, ci] += params.PREVIOUS
                for advisor in advisors:
                    if advisor in instructors:
                        weights[si, ci] += params.ADVISORS
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
            rank_rows.append(pd.Series(collect, name=row['NetID']))

    df = pd.concat(rank_rows, axis='columns').T
    df['Bank'] = pd.to_numeric(df['Bank'], errors='coerce', downcast='float')
    df['Join'] = pd.to_numeric(df['Join'], errors='coerce', downcast='float')
    df['Weight'] = pd.to_numeric(
        df['Weight'], errors='coerce', downcast='float').fillna(
        params.DEFAULT_ASSIGN)
    df['Weight'] += params.BANK_MULTIPLIER * (
            df['Bank'].fillna(params.DEFAULT_BANK) - params.DEFAULT_BANK)
    df['Weight'] += params.JOIN_MULTIPLIER * (
            df['Join'].fillna(params.DEFAULT_JOIN) - params.DEFAULT_JOIN)
    df['Weight'] += params.MSE_BOOST * (
            (df['Year'] == 'MSE1') + (df['Year'] == 'MSE2'))
    return df


def read_course_data(filename: str) -> pd.DataFrame:
    rank_rows = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {col: row[col] for col in course_data_cols}
            # expand list of students into ranking for each student
            for rank in course_options:
                for student in row[rank].split(';'):
                    if student:
                        collect[student] = rank
            rank_rows.append(
                rank_rows.append(pd.Series(collect, name=row['Course'])))

    df = pd.concat(rank_rows, axis='columns').T
    df['Slots'] = pd.to_numeric(
        df['Slots'], errors='coerce', downcast='integer').fillna(1)
    df['First weight'] = pd.to_numeric(
        df['Weight'], errors='coerce', downcast='float').fillna(
        params.DEFAULT_FIRST_FILL)
    df['Base weight'] = params.DEFAULT_BASE_FILL
    return df


def check_input(student_data: pd.Series, course_data: pd.Series):
    for index in student_data[student_data.index.duplicated()].index:
        sys.exit(
            'Duplicate rows for netid {}. Exiting without solving'.format(
                index))
    for index in course_data[course_data.index.duplicated()].index:
        sys.exit(
            'Duplicate rows for course {}. Exiting without solving'.format(
                index))


def matching_differences(extra_course: str, old: List[Tuple[int, int]],
                         new: List[Tuple[int, int]], student_data: pd.DataFrame,
                         course_data: pd.DataFrame,
                         student_scores: pd.DataFrame,
                         course_scores: pd.DataFrame):
    df_s = pd.concat(
        [pd.Series(dict(old), dtype=int), pd.Series(dict(new), dtype=int)],
        axis=1)
    student_diffs = df_s.iloc[:, 0].compare(df_s.iloc[:, 1]).astype(int)
    course_diffs, student_changes = get_matching_changes(
        course_data, extra_course, student_data, student_diffs, student_scores)
    course_changes = parse_course_changes(
        course_data, course_diffs, course_scores)

    return student_changes, course_changes


def parse_course_changes(course_data, course_diffs, course_scores):
    course_changes = []
    for _, row in course_diffs.iterrows():
        students_old = str(row[0]).split(';')
        students_new = str(row[1]).split(';')
        for old, new in zip(students_old, students_new):
            score_change = 'NA'
            if old in course_data.columns and new in course_data.columns:
                score_change = course_scores.loc[row.name, new] - \
                               course_scores.loc[row.name, old]

            if old in course_data.columns:
                old += ' ({})'.format(course_data.loc[row.name, old])
            if new in course_data.columns:
                new += ' ({})'.format(course_data.loc[row.name, new])
            course_changes.append((row.name, old, new, score_change))
    return course_changes


def get_matching_changes(course_data, extra_course, student_data, student_diffs,
                         student_scores):
    course_matches_old = defaultdict(list)
    course_matches_new = defaultdict(list)
    student_changes = []
    for _, row in student_diffs.iterrows():
        student = student_data.index[row.name]
        course_old = course_data.index[row[0]] if row[0] >= 0 else 'unassigned'
        course_new = course_data.index[row[1]] if row[1] >= 0 else 'unassigned'
        score_change = 'NA'
        if row[0] >= 0 and row[1] >= 0:
            score_change = student_scores.loc[student, course_new] - \
                           student_scores.loc[student, course_old]

        if row[0] >= 0:
            course_matches_old[course_old].append(student)
            course_old += ' ({})'.format(student_data.loc[student, course_old])
        if row[1] >= 0:
            course_matches_new[course_new].append(student)
            course_new += ' ({})'.format(student_data.loc[student, course_new])

        student_changes.append((student, course_old, course_new, score_change))
    if extra_course:
        course_matches_new[extra_course].append('extra')
    course_matches_old = {course: ';'.join(sorted(students)) for
                          course, students in course_matches_old.items()}
    course_matches_new = {course: ';'.join(sorted(students)) for
                          course, students in course_matches_new.items()}
    df_c = pd.concat(
        [pd.Series(course_matches_old, dtype=str),
         pd.Series(course_matches_new, dtype=str)], axis=1).fillna('none')
    course_diffs = df_c.iloc[:, 0].compare(df_c.iloc[:, 1])
    return course_diffs, student_changes


def single_line(changes):
    return ', '.join(
        ['{}: {} -> {}'.format(agent, old, new) for agent, old, new, _ in
         changes])


def test_additional_TA(path, student_data, course_data, student_scores,
                       course_scores, weights, fixed_matches, graph):
    data = {}
    initial_matches = graph.get_matching(fixed_matches)
    for ci, course in enumerate(course_data.index):
        # fill one additional course slot
        fixed_edit = pd.concat(
            [fixed_matches.T, pd.Series(
                ['', course, -1, ci],
                index=['Student', 'Course', 'Student index', 'Course index'])],
            axis=1).T
        graph_edit = min_cost_flow.MatchingGraph(
            weights, student_data['Weight'],
            course_data[['Slots', 'Base weight', 'First weight']], fixed_edit)
        if graph_edit.solve():
            new_matches = graph_edit.get_matching(fixed_edit)
            student_changes, course_changes = matching_differences(
                course, initial_matches, new_matches, student_data, course_data,
                student_scores, course_scores)
            data[course] = (
                                   graph.flow.OptimalCost() - graph_edit.flow.OptimalCost()) / 10 ** min_cost_flow.DIGITS, single_line(
                student_changes), single_line(course_changes)
        else:
            data[course] = -100, '', ''

    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(
            ['Course', 'Weight change', 'Student differences',
             'Course differences'])
        for course, tup in sorted(data.items(), key=lambda item: -item[1][0]):
            writer.writerow(
                [course, tup[0] if tup[0] != -100 else 'Error', tup[1], tup[2]])


def test_removing_TA(path, student_data, course_data, student_scores,
                     course_scores, weights, fixed_matches, graph):
    data = {}
    initial_matches = graph.get_matching(fixed_matches)
    for si, student in enumerate(student_data.index):
        bank = str(student_data.loc[student, 'Bank']).replace('nan', '')
        join = str(student_data.loc[student, 'Join']).replace('nan', '')
        if si not in fixed_matches['Student index'].values:
            # no edges from student node
            fixed_edit = pd.concat(
                [fixed_matches.T, pd.Series(
                    [student, '', si, -1],
                    index=['Student', 'Course', 'Student index',
                           'Course index'])], axis=1).T
            graph_edit = min_cost_flow.MatchingGraph(
                weights, student_data['Weight'],
                course_data[['Slots', 'Base weight', 'First weight']],
                fixed_edit)
            if graph_edit.solve():
                new_matches = graph_edit.get_matching(fixed_edit)
                student_changes, course_changes = matching_differences(
                    None, initial_matches, new_matches, student_data,
                    course_data, student_scores, course_scores)
                weight_change = (
                                        graph.flow.OptimalCost() - graph_edit.flow.OptimalCost()) / 10 ** min_cost_flow.DIGITS
                data[student] = weight_change, bank, join, single_line(
                    student_changes), single_line(
                    course_changes)
            else:
                data[student] = -100, bank, join, '', ''

    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(
            ['Student', 'Weight change', 'Bank', 'Join', 'Student differences',
             'Course differences'])
        for student, tup in sorted(data.items(), key=lambda item: -item[1][0]):
            writer.writerow(
                [student, tup[0] if tup[0] != -100 else 'Error', tup[1], tup[2],
                 tup[3], tup[4]])


def find_alternate_matching(path, student_data, course_data, weights,
                            fixed_matches, initial_matches, last_matches,
                            cumulative, student_scores, course_scores,
                            step=0.1):
    while True:
        for si, ci in initial_matches:
            if ci >= 0:
                weights[si, ci] -= step
        cumulative += step
        graph_edit = min_cost_flow.MatchingGraph(
            weights, student_data['Weight'],
            course_data[['Slots', 'Base weight', 'First weight']],
            fixed_matches)
        if graph_edit.solve():
            new_matches = graph_edit.get_matching(fixed_matches)
            student_changes, course_changes = matching_differences(
                None, last_matches, new_matches, student_data, course_data,
                student_scores, course_scores)
            if len(student_changes) > 0 or len(course_changes) > 0:
                student_changes, course_changes = matching_differences(
                    None, initial_matches, new_matches, student_data,
                    course_data, student_scores, course_scores)
                new_weight = -graph_edit.flow.OptimalCost() / 100 + (
                        len(new_matches) - len(student_changes)) * cumulative
                print(
                    'Solved alternate flow with total weight {:.2f}'.format(
                        new_weight))
                with open(path, 'w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(
                        ['Student/Course', 'Previous match', 'New match',
                         'Preference change'])
                    for student, old, new, score_change in student_changes:
                        writer.writerow([student, old, new, score_change])
                    for course, old, new, score_change in course_changes:
                        writer.writerow([course, old, new, score_change])
                return new_matches, cumulative, new_weight


def write_params(output_path):
    with open(output_path + 'params.csv', 'w') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'Weight'])
        for param, param_val in params.__dict__.items():
            if param.startswith("__"): continue
            writer.writerow([param, param_val])


def run_matching(path="", student_data="inputs/student_data.csv",
                 course_data="inputs/course_data.csv", fixed="inputs/fixed.csv",
                 adjusted="inputs/adjusted.csv", output="outputs/",
                 alternates=2) -> Tuple[float, int, List[float]]:
    path = validate_path_args(path, output)
    student_data, course_data = read_student_and_course_data(
        path, student_data, course_data)

    student_scores = get_student_scores(student_data, course_data.index)
    course_scores = get_course_scores(course_data, student_data.index)

    weights = match_weights(
        student_data, student_scores, course_data, course_scores)

    make_manual_adjustments(
        course_scores, path + adjusted, student_scores, weights)

    fixed_matches = get_fixed_matches(
        course_scores, fixed, path, student_scores)

    graph = min_cost_flow.MatchingGraph(
        weights, student_data['Weight'],
        course_data[['Slots', 'Base weight', 'First weight']], fixed_matches)

    matching_weight = 0
    slots_unfilled = float('inf')
    if graph.solve():
        matching_weight, slots_unfilled = graph.write_matching(
            path + output + 'matchings.csv', weights, student_data,
            student_scores, course_data, course_scores,
            fixed_matches[['Student index', 'Course index']])
        print(f'Solved optimal flow with total weight {matching_weight:.2f}')

        write_params(path + output)

    else:
        print('Problem optimizing flow')
        graph.print()

    test_additional_TA(
        path + output + 'additional_TA.csv', student_data, course_data,
        student_scores, course_scores, weights, fixed_matches, graph)
    test_removing_TA(
        path + output + 'remove_TA.csv', student_data, course_data,
        student_scores, course_scores, weights, fixed_matches, graph)

    alt_weights = run_alternate_matchings(
        alternates, course_data, course_scores, fixed_matches, graph, output,
        path, student_data, student_scores, weights)

    return matching_weight, slots_unfilled, alt_weights


def make_manual_adjustments(course_scores, path_adjusted, student_scores,
                            weights):
    if not os.path.isfile(path_adjusted): return
    adjusted_matches = pd.read_csv(
        path_adjusted, dtype={'NetID': str, 'Course': str, 'Weight': float})
    for _, row in adjusted_matches.iterrows():
        si = student_scores.index.get_loc(row['NetID'])
        if row['Course'] not in course_scores.index:
            continue
        ci = course_scores.index.get_loc(row['Course'])
        weights[si, ci] += row['Weight']


def read_student_and_course_data(path, student_data, course_data):
    student_data = read_student_data(path + student_data)
    course_data = read_course_data(path + course_data)
    for student in student_data.index:
        if student not in course_data.columns.values:
            course_data[student] = np.nan
    for course in course_data.index:
        if course not in student_data.columns.values:
            student_data[course] = np.nan
    return student_data, course_data


def get_fixed_matches(course_scores, fixed, path, student_scores):
    if os.path.isfile(path + fixed):
        fixed_matches = pd.read_csv(path + fixed, dtype=str)
        fixed_matches['Student index'] = [student_scores.index.get_loc(student)
                                          for student in fixed_matches['NetID']]
        fixed_matches['Course index'] = [course_scores.index.get_loc(course) for
                                         course in fixed_matches['Course']]
    else:
        fixed_matches = pd.DataFrame(
            columns=['NetID', 'Course', 'Student index', 'Course index'])
    return fixed_matches


def run_alternate_matchings(alternates, course_data, course_scores,
                            fixed_matches, graph, output, path, student_data,
                            student_scores, weights):
    best_matches = graph.get_matching(fixed_matches)
    last_matches = best_matches
    cumulative = 0
    alt_weights = []
    for i in range(alternates):
        last_matches, cumulative, alt_weight = find_alternate_matching(
            path + output + 'alternate{}.csv'.format(i + 1), student_data,
            course_data, weights, fixed_matches, best_matches, last_matches,
            cumulative, student_scores, course_scores)
        alt_weights.append(alt_weight)
    return alt_weights


def validate_path_args(path, output):
    # ensure trailing slash on path directory
    if path[-1] != '/':
        path += '/'

    # ensure output directory exists
    os.makedirs(path + output, exist_ok=True)
    return path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Generate matching from input data.')
    parser.add_argument(
        '--path', metavar='FOLDER', default='.', help='prefix to file paths')
    parser.add_argument(
        '--student_data', metavar='STUDENT DATA',
        default='inputs/student_data.csv',
        help='csv file with student data rows')
    parser.add_argument(
        '--course_data', metavar='COURSE DATA',
        default='inputs/course_data.csv', help='csv file with course data rows')
    parser.add_argument(
        '--fixed', metavar='FIXED INPUT', default='inputs/fixed.csv',
        help='csv file with fixed student-course matchings')
    parser.add_argument(
        '--adjusted', metavar='ADJUSTED INPUT', default='inputs/adjusted.csv',
        help='csv file with adjustment weights for student-course matchings')
    parser.add_argument(
        '--output', metavar='MATCHING OUTPUT', default='outputs/',
        help='location to write matching output')
    parser.add_argument(
        '--alternates', metavar='ALTERNATE MATCHES', type=int, default=3,
        help='number of alternate matches to solve for')
    args = parser.parse_args()
    args.path = validate_path_args(args.path, args.output)

    run_matching(**vars(args))
