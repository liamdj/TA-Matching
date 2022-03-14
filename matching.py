import argparse
import csv
import math
import os
import sys
from collections import defaultdict
from typing import Any, List, Tuple, Optional, DefaultDict, Dict

import numpy as np
import pandas as pd

import min_cost_flow
import params


def __is_nan(num: Any) -> bool:
    if type(num) == tuple:
        return False
    if type(num) == str:
        return num == 'nan'
    return math.isnan(num) or pd.isna(num)


student_data_cols = ['Name', 'Weight', 'Year', 'Bank', 'Join', 'Previous',
                     'Advisors', 'Notes']
student_options = ['Okay', 'Good', 'Favorite']
course_data_cols = ['Slots', 'Weight', 'Instructor']
course_options = ['Veto', 'Favorite']

ChangeDetails = List[Tuple[str, str, str]]


def match_weights(student_data: pd.DataFrame, course_data: pd.DataFrame,
                  adjusted_path: str, default_value: float = np.nan,
                  ignore_instructor_prefs=False) -> np.ndarray:
    def get_rank(rank) -> Tuple[str, int]:
        if ignore_instructor_prefs or __is_nan(rank):
            return np.nan, 0
        return rank

    weights = np.full(
        (len(student_data.index), len(student_data.index)), default_value)
    for si, (student, s_data) in enumerate(student_data.iterrows()):
        previous = s_data['Previous'].split(';')
        advisors = s_data['Advisors'].split(';')
        for ci, (course, c_data) in enumerate(course_data.iterrows()):
            instructors = c_data['Instructor'].split(';')
            c_rank = get_rank(c_data[student])
            s_rank = get_rank(s_data[course])
            if c_rank[0] != 'Veto' and course in s_data and not __is_nan(
                    s_rank[0]):
                weights[si, ci] = calculate_weight_per_pairing(
                    course, s_rank, c_rank, instructors, advisors,
                    previous, s_data['Ranked'], c_data['Favorites'])
    make_manual_adjustments(student_data, course_data, weights, adjusted_path)
    return weights


def calculate_weight_per_pairing(course: str, s_rank: Tuple[str, int],
                                 c_rank: Tuple[str, int],
                                 instructors: List[str], advisors: List[str],
                                 previous: List[str], s_ranked: int,
                                 c_favorites: int) -> float:
    weight = (params.BOOST_PER_COURSE_STUDENT_RANKED * s_ranked) + (
            params.BOOST_PER_FAVORITE_STUDENT * c_favorites)
    c_sorted_boost = c_rank[1] * params.BOOST_PER_PLACE_IN_SORTED_COURSE_LIST
    s_sorted_boost = s_rank[1] * params.BOOST_PER_PLACE_IN_SORTED_STUDENT_LIST
    if s_rank[0] == 'Favorite':
        if c_rank[0] == 'Favorite':
            weight += params.FAVORITE_FAVORITE + c_sorted_boost + s_sorted_boost
        else:
            weight += params.STUDENT_FAVORITE_INSTRUCTOR_NEUTRAL + s_sorted_boost
    elif s_rank[0] == 'Good' and c_rank[0] == 'Favorite':
        weight += params.STUDENT_GOOD_INSTRUCTOR_FAVORITE + c_sorted_boost
    elif s_rank[0] == 'Okay':
        weight += params.OKAY_COURSE_PENALTY

    if course in previous:
        weight += params.PREVIOUS
    for advisor in advisors:
        if advisor in instructors:
            weight += params.ADVISORS
    return weight


def read_student_data(filename: str) -> pd.DataFrame:
    """
    Returns a dataframe with rows indexed by NetIDs and columns specified by
    `student_data_cols` _and_ every course (with the corresponding values being
    `Okay`, `Good`, `Favorite`, or `NaN` (if no value was specified). Also,
    there is the field 'Ranked' that is the # of `Favorite` and `Good` courses
    enumerated by each student.
    """
    rank_rows = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {col: row[col] for col in student_data_cols}
            # expand list of courses into ranking for each course
            collect['Ranked'] = 0

            sorted_favs = row.get('Sorted Favorites', '').split('>')
            leveled_favs = {}
            for i, fav_level in enumerate(sorted_favs):
                for fav in fav_level.split('='):
                    if fav:
                        leveled_favs[fav] = len(sorted_favs) - i

            for rank in student_options:
                courses_in_rank = list(
                    filter(
                        lambda x: True if x else False, row[rank].split(';')))
                if rank == 'Favorite' or rank == 'Good':
                    collect['Ranked'] += len(courses_in_rank)
                for i, course in enumerate(courses_in_rank):
                    sorted_score = 0  # implies in no order
                    if rank == 'Favorite':
                        sorted_score = leveled_favs.get(course, 0)
                    collect[course] = (rank, sorted_score)
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
    """
    Returns a dataframe with rows indexed by course names and columns specified
    by `course_data_cols` _and_ every student (with the corresponding values being
    being `Veto`, `Favorite`, or `NaN` (if no value was specified).
    """
    rank_rows = []
    with open(filename, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            collect = {col: row[col] for col in course_data_cols}
            # expand list of students into ranking for each student
            for rank in course_options:
                courses_in_rank = list(
                    filter(
                        lambda x: True if x and not x.isspace() else False,
                        row[rank].split(';')))
                if rank == 'Favorite':
                    collect['Favorites'] = len(courses_in_rank)
                for i, student in enumerate(courses_in_rank):
                    sorted_score = 0  # implies in no order
                    if rank == 'Favorite' and row.get('Favorites Sorted'):
                        sorted_score = len(courses_in_rank) - i  # positive
                    collect[student] = (rank, sorted_score)
            rank_rows.append(pd.Series(collect, name=row['Course']))

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
        sys.exit(f'Duplicate rows for netid {index}. Exiting without solving')
    for index in course_data[course_data.index.duplicated()].index:
        sys.exit(f'Duplicate rows for course {index}. Exiting without solving')


def matching_differences(extra_course: Optional[str],
                         old: List[Tuple[int, int]], new: List[Tuple[int, int]],
                         student_data: pd.DataFrame,
                         course_data: pd.DataFrame) -> Tuple[
    ChangeDetails, ChangeDetails]:
    df_s = pd.concat(
        [pd.Series(dict(old), dtype=int), pd.Series(dict(new), dtype=int)],
        axis=1)
    student_diffs = df_s.iloc[:, 0].compare(df_s.iloc[:, 1]).astype(int)
    course_diffs, student_changes = get_matching_changes(
        course_data, extra_course, student_data, student_diffs)
    course_changes = parse_course_changes(course_data, course_diffs)
    student_changes.sort(key=lambda x: x[0])
    course_changes.sort(key=lambda x: x[0])
    return student_changes, course_changes


def parse_course_changes(course_data: pd.DataFrame,
                         course_diffs: pd.DataFrame) -> ChangeDetails:
    course_changes = []
    for _, row in course_diffs.iterrows():
        students_old = str(row[0]).split(';')
        students_new = str(row[1]).split(';')
        for old, new in zip(students_old, students_new):
            if old in course_data.columns:
                old += ' ({})'.format(course_data.loc[row.name, old])
            if new in course_data.columns:
                new += ' ({})'.format(course_data.loc[row.name, new])
            course_changes.append((str(row.name), old, new))
    return course_changes


def get_matching_changes(course_data: pd.DataFrame, extra_course: Optional[str],
                         student_data: pd.DataFrame,
                         student_diffs: pd.DataFrame) -> Tuple[
    pd.DataFrame, ChangeDetails]:
    course_matches_old = defaultdict(list)
    course_matches_new = defaultdict(list)
    student_changes = parse_student_changes(
        course_data, course_matches_new, course_matches_old, student_data,
        student_diffs)
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


def parse_student_changes(course_data: pd.DataFrame,
                          course_matches_new: DefaultDict,
                          course_matches_old: DefaultDict,
                          student_data: pd.DataFrame,
                          student_diffs: pd.DataFrame) -> ChangeDetails:
    student_changes = []
    for _, row in student_diffs.iterrows():
        student = student_data.index[row.name]
        course_old = course_data.index[row[0]] if row[0] >= 0 else 'unassigned'
        course_new = course_data.index[row[1]] if row[1] >= 0 else 'unassigned'
        if row[0] >= 0:
            course_matches_old[course_old].append(student)
            course_old += ' ({})'.format(student_data.loc[student, course_old])
        if row[1] >= 0:
            course_matches_new[course_new].append(student)
            course_new += ' ({})'.format(student_data.loc[student, course_new])

        student_changes.append((student, course_old, course_new))
    return student_changes


def single_line(changes: ChangeDetails) -> str:
    return ',\n'.join(
        ['{}: {} -> {}'.format(agent, old, new) for agent, old, new in changes])


def test_additional_TA(path: str, student_data: pd.DataFrame,
                       course_data: pd.DataFrame, weights: np.ndarray,
                       fixed_matches: pd.DataFrame,
                       initial_matches: List[Tuple[int, int]],
                       initial_matching_weight: float):
    data = {}
    for ci, course in enumerate(course_data.index):
        # fill one additional course slot
        fixed_edit = pd.concat(
            [fixed_matches.T, pd.Series(
                [np.nan, course, -1, ci],
                index=['Student', 'Course', 'Student index', 'Course index'])],
            axis=1).T
        data[course] = calculate_changes_in_new_graph(
            student_data, course_data, initial_matches, weights, fixed_edit,
            initial_matching_weight, course)
    write_edited_graph_changes(
        path, data, ['Course', 'Weight change', 'Student differences',
                     'Course differences'])


def write_edited_graph_changes(path: str, change_data: Dict[str, Tuple],
                               columns: List[str]):
    sorted_by_weight = sorted(change_data.items(), key=lambda item: -item[1][0])
    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(columns)
        for change_index, (weight_change, *tup) in sorted_by_weight:
            weight_change = round(
                weight_change, 4) if weight_change != -100 else 'Error'
            writer.writerow([change_index, weight_change, *tup])


def calculate_changes_in_new_graph(student_data: pd.DataFrame,
                                   course_data: pd.DataFrame,
                                   initial_matches: List[Tuple[int, int]],
                                   weights: np.ndarray,
                                   fixed_matches: pd.DataFrame,
                                   old_weight: float,
                                   extra_course: str = None) -> Tuple[
    float, str, str]:
    changes = make_changes_and_calculate_differences(
        student_data, course_data, initial_matches, weights, fixed_matches,
        extra_course)
    if not changes:
        return -100, '', ''

    student_changes, course_changes, new_weight = changes
    w_change = new_weight - old_weight
    return w_change, single_line(student_changes), single_line(course_changes)


def make_changes_and_calculate_differences(student_data: pd.DataFrame,
                                           course_data: pd.DataFrame,
                                           initial_matches: List[
                                               Tuple[int, int]],
                                           weights: np.ndarray,
                                           fixed_matches: pd.DataFrame,
                                           extra_course: str = None) -> \
        Optional[Tuple[ChangeDetails, ChangeDetails, float]]:
    """ Only returns `None` if the graph could not be solved. """
    graph = min_cost_flow.MatchingGraph(
        weights, student_data['Weight'],
        course_data[['Slots', 'Base weight', 'First weight']], fixed_matches)

    if not graph.solve():
        return None
    new_matches = graph.get_matching(fixed_matches, weights)
    student_changes, course_changes = matching_differences(
        extra_course, initial_matches, new_matches, student_data, course_data)
    return student_changes, course_changes, graph.graph_weight()


def test_removing_TA(path: str, student_data: pd.DataFrame,
                     course_data: pd.DataFrame, weights: np.ndarray,
                     fixed_matches: pd.DataFrame,
                     initial_matches: List[Tuple[int, int]],
                     initial_matching_weight: float):
    data = {}
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
            weight_change, s_changes, c_changes = calculate_changes_in_new_graph(
                student_data, course_data, initial_matches, weights, fixed_edit,
                initial_matching_weight, None)
            data[student] = weight_change, bank, join, s_changes, c_changes

    write_edited_graph_changes(
        path, data,
        ['Student', 'Weight change', 'Bank', 'Join', 'Student differences',
         'Course differences'])


def test_adding_or_subtracting_a_slot(path: str, add: bool,
                                      student_data: pd.DataFrame,
                                      course_data: pd.DataFrame,
                                      weights: np.ndarray,
                                      fixed_matches: pd.DataFrame,
                                      initial_matches: List[Tuple[int, int]],
                                      initial_match_weight: float):
    """ if `add == True`, add a slot, otherwise subtract a slot """
    data = {}
    s_d = (2 * add) - 1
    for ci, course in enumerate(course_data.index):
        course_data.at[course, 'Slots'] = course_data.loc[course, 'Slots'] + s_d
        data[course] = calculate_changes_in_new_graph(
            student_data, course_data, initial_matches, weights, fixed_matches,
            initial_match_weight, course)
        course_data.at[course, 'Slots'] = course_data.loc[course, 'Slots'] - s_d
    write_edited_graph_changes(
        path, data, ['Course', 'Weight change', 'Student differences',
                     'Course differences'])


def find_alternate_matching(path: str, student_data: pd.DataFrame,
                            course_data: pd.DataFrame, weights: np.ndarray,
                            fixed_matches: pd.DataFrame,
                            initial_matches: List[Tuple[int, int]],
                            last_matches: List[Tuple[int, int]],
                            cumulative: float, step=0.1) -> Tuple[
    List[Tuple[int, int]], float, float]:
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
            new_matches = graph_edit.get_matching(fixed_matches, weights)
            student_changes, course_changes = matching_differences(
                None, last_matches, new_matches, student_data, course_data)
            if len(student_changes) > 0 or len(course_changes) > 0:
                new_weight = write_alternate_match(
                    path, graph_edit.graph_weight(), student_data, course_data,
                    initial_matches, new_matches, cumulative)
                return new_matches, cumulative, new_weight


def write_alternate_match(path: str, graph_weight: float,
                          student_data: pd.DataFrame, course_data: pd.DataFrame,
                          initial_matches: List[Tuple[int, int]],
                          new_matches: List[Tuple[int, int]],
                          cumulative: float) -> float:
    student_changes, course_changes = matching_differences(
        None, initial_matches, new_matches, student_data, course_data)
    new_weight = graph_weight + (
            len(new_matches) - len(student_changes)) * cumulative
    print(f'Solved alternate flow with total weight {new_weight:.2f}')
    with open(path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(
            ['Student/Course', 'Previous match', 'New match'])
        for student, old, new in student_changes:
            writer.writerow([student, old, new])
        for course, old, new in course_changes:
            writer.writerow([course, old, new])
    return new_weight


def write_params(output_path: str):
    with open(output_path + 'params.csv', 'w') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'Weight'])
        for param, param_val in params.__dict__.items():
            if param.startswith("__"):
                continue
            writer.writerow([param, param_val])


def run_matching(path="", student_data="inputs/student_data.csv",
                 course_data="inputs/course_data.csv", fixed="inputs/fixed.csv",
                 adjusted="inputs/adjusted.csv", previous="inputs/previous.csv",
                 output="outputs/", alternates=2, run_interviews=False) -> \
        Tuple[float, int, List[float]]:
    path = validate_path_args(path, output)
    student_data, course_data = read_student_and_course_data(
        path, student_data, course_data)
    weights = match_weights(student_data, course_data, path + adjusted)
    previous_matches = make_adjustments_from_previous(
        course_data, path + previous, student_data, weights)
    fixed_matches = get_fixed_matches(course_data, path + fixed, student_data)
    graph = min_cost_flow.MatchingGraph(
        weights, student_data['Weight'],
        course_data[['Slots', 'Base weight', 'First weight']], fixed_matches)

    if not graph.solve():
        print('Problem optimizing flow')
        graph.print()
        return -1.0, course_data['Slots'].sum(), []

    output_path = path + output
    matching_weight, slots_unfilled, initial_matches = graph.write_matching(
        output_path + 'matchings.csv', weights, student_data, course_data,
        fixed_matches)
    print(f'Solved optimal flow with total weight {matching_weight:.2f}')

    write_params(output_path)

    alt_weights = run_additional_features(
        output_path, student_data, course_data, weights, fixed_matches,
        initial_matches, matching_weight, alternates, previous_matches,
        path + adjusted, run_interviews)

    return matching_weight, slots_unfilled, alt_weights


def run_additional_features(output_path: str, student_data: pd.DataFrame,
                            course_data: pd.DataFrame, weights: np.ndarray,
                            fixed_matches: pd.DataFrame,
                            initial_matches: List[Tuple[int, int]],
                            matching_weight: float, alternates: int,
                            previous_matches: pd.DataFrame, adjusted_path: str,
                            run_interviews=False) -> List[float]:
    def repopulate_weights(value: float):
        for si, ci in initial_matches:
            weights[si, ci] += value

    repopulate_weights(params.PREVIOUS_MATCHING_BOOST)
    test_additional_TA(
        output_path + 'additional_TA.csv', student_data, course_data, weights,
        fixed_matches, initial_matches, matching_weight)
    test_removing_TA(
        output_path + 'remove_TA.csv', student_data, course_data, weights,
        fixed_matches, initial_matches, matching_weight)
    test_adding_or_subtracting_a_slot(
        output_path + 'add_slot.csv', True, student_data, course_data, weights,
        fixed_matches, initial_matches, matching_weight)
    test_adding_or_subtracting_a_slot(
        output_path + 'remove_slot.csv', False, student_data, course_data,
        weights, fixed_matches, initial_matches, matching_weight)
    repopulate_weights(-params.PREVIOUS_MATCHING_BOOST)

    alt_weights = run_alternate_matchings(
        alternates, course_data, fixed_matches, initial_matches, output_path,
        student_data, weights)
    test_changes_from_previous(
        output_path, previous_matches, student_data, course_data, weights,
        fixed_matches)
    if run_interviews:
        interview_list(
            course_data, student_data, fixed_matches, adjusted_path,
            output_path, initial_matches)
    return alt_weights


def test_changes_from_previous(output_path: str, previous_matches: pd.DataFrame,
                               student_data: pd.DataFrame,
                               course_data: pd.DataFrame, weights: np.ndarray,
                               fixed_matches: pd.DataFrame):
    def get_student_and_course_indices(student: pd.Series) -> Tuple[int, int]:
        netid = student["NetID"]
        course = student["Course"]
        si, ci = -1, -1
        if netid in student_data.index:
            si = student_data.index.get_loc(netid)
        if course in course_data.index:
            ci = course_data.index.get_loc(course)
        return si, ci

    def repopulate_weights(value: float):
        for _, entry in previous_matches.iterrows():
            si, ci = get_student_and_course_indices(entry)
            if si != -1 and ci != -1:
                weights[si, ci] += value

    def binary_search(desired_matches: int,
                      initial_matches: List[Tuple[int, int]]) -> Optional[
        Tuple[int, ChangeDetails, ChangeDetails]]:
        lo = 0.0
        hi = 50.0
        prev_desired = None
        while hi > lo + 0.01:
            mid = lo + (hi - lo) / 2
            repopulate_weights(mid)
            changes = make_changes_and_calculate_differences(
                student_data, course_data, initial_matches, weights,
                fixed_matches)
            repopulate_weights(-mid)
            if not changes or len(changes[0]) > desired_matches:
                lo = mid
            else:
                hi = mid
                if len(changes[0]) == desired_matches:
                    prev_desired = (mid, *changes)
        return None if not prev_desired else prev_desired[:3]

    def get_previous_indices():
        indices = []
        for _, match in previous_matches.iterrows():
            si, ci = get_student_and_course_indices(match)
            if si != -1:
                indices.append((si, ci))
        return indices

    def get_initial_changes() -> Tuple[int, str, str]:
        changes = make_changes_and_calculate_differences(
            student_data, course_data, previous_indices, weights, fixed_matches)
        if not changes:
            print(f"Graph could not be solved with no added weight")
            return -1, "", ""
        return len(changes[0]), single_line(changes[0]), single_line(
            changes[1])

    # remove the weight that was added earlier to boost previous matches
    repopulate_weights(-params.PREVIOUS_MATCHING_BOOST)
    previous_indices = get_previous_indices()

    max_changes, orig_student_diffs, orig_course_diffs = get_initial_changes()
    if max_changes == -1:
        return
    found_changes = [(max_changes, 0.0, orig_student_diffs, orig_course_diffs)]

    for i in range(1, max_changes):
        change = binary_search(i, previous_indices)
        if change:
            weight_added_per_prev_match, student_diffs, course_diffs = change
            if len(student_diffs) == i:
                found_changes.append(
                    (i, round(weight_added_per_prev_match, 4),
                     single_line(student_diffs), single_line(course_diffs)))
            else:
                print(f"Problem in trying to get exactly {i} student changes")
        else:
            found_changes.append((i, 'Impossible', '', ''))

    with open(output_path + 'weighted_changes.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerow(
            ['Desired Student Changes', 'Weight Added to Previous Match',
             'Student Changes', 'Course Changes'])
        writer.writerows(sorted(found_changes, key=lambda x: x[0]))


def calculate_l2norm_for_uniform_entries(matched_students: int,
                                         denominator: int) -> float:
    """
    Returns the L2 (Frobenius) norm for a matrix that would have `denominator`
    entries of `1/denominator` in each row, with `matched_students` number of
    non-zero rows.
    """
    return math.sqrt(matched_students * 1 / denominator)


def add_noise_and_get_matching_interviews(stddev: float,
                                          base_weights: np.ndarray,
                                          student_weight_data: pd.DataFrame,
                                          course_graph_data: pd.DataFrame,
                                          fixed_matches: pd.DataFrame) -> Tuple[
    List[Tuple[int, int]], int]:
    noise = np.random.normal(0, stddev, base_weights.shape)
    new_weights = base_weights + noise
    graph = min_cost_flow.MatchingGraph(
        new_weights, student_weight_data, course_graph_data, fixed_matches)
    if not graph.solve():
        print('Problem optimizing flow')
        graph.print()
        return [], -1
    match = graph.get_matching(fixed_matches, base_weights)
    unfilled_slots = graph.get_slots_unfilled(graph.get_slots_filled(match))
    return match, unfilled_slots


def run_interview_trials(sigma: float, trials_to_run: int,
                         base_weights: np.ndarray,
                         student_weight_data: pd.DataFrame,
                         course_graph_data: pd.DataFrame,
                         fixed_matches: pd.DataFrame) -> np.ndarray:
    courses = len(course_graph_data.index)
    students = len(student_weight_data.index)
    trial = np.zeros((students, courses), dtype=np.float64)
    for _ in range(trials_to_run):
        matching, unfilled = add_noise_and_get_matching_interviews(
            sigma, base_weights, student_weight_data, course_graph_data,
            fixed_matches)
        if unfilled == 0:
            for si, ci in matching:
                if si != -1 and ci != -1:
                    trial[si, ci] += 1.0
    return trial


def check_caps_interview_simulations(course_data: pd.DataFrame,
                                     recent_simulation: np.ndarray,
                                     batch_num: int,
                                     num_trials: int,
                                     previous_simulations_percentages: np.ndarray) -> bool:
    courses = len(course_data.index)
    comparisons = np.absolute(
        (recent_simulation / num_trials) - previous_simulations_percentages)
    hard_cap = (comparisons >= 0.05).sum()
    if hard_cap == 0:
        soft_cap = (comparisons >= 0.02).sum(axis=0)
        courses_above_soft_cap = []
        for ci in range(courses):
            ci_slots = course_data.iloc[ci]['Slots']
            if soft_cap[ci] > int(ci_slots * 0.45):
                courses_above_soft_cap.append(
                    (course_data.index[ci], soft_cap[ci], ci_slots))

        if len(courses_above_soft_cap) == 0:
            print(
                f"Stopping after batch {batch_num} with {num_trials} trials due to being under the soft cap")
            return True
        else:
            print(
                f"After batch {batch_num} with {num_trials} trials, above soft cap ('Course', 'TAs > 2%', 'Slots'): {courses_above_soft_cap}")
    else:
        print(
            f"After batch {batch_num} with {num_trials} trials, still {hard_cap} differences above 0.05")
    return False


def interview_list(course_data: pd.DataFrame, student_data: pd.DataFrame,
                   fixed_matches: pd.DataFrame, adjusted_path: str,
                   output_path: str, initial_matches: List[Tuple[int, int]]):
    courses = len(course_data.index)
    students = len(student_data.index)
    student_weight_data = student_data['Weight']
    course_graph_data = course_data[['Slots', 'Base weight', 'First weight']]
    cumulative_percentages = np.zeros((students, courses), dtype=np.float64)
    for si, ci in initial_matches:
        cumulative_percentages[si, ci] += 1.0
    weights = match_weights(student_data, course_data, adjusted_path, 0.0, True)
    max_stddev, step = 3.5, 0.5
    simulations = 1
    for sigma in np.arange(step, max_stddev + step, step):
        trials = 50
        sim_trials = trials
        sim_matches = run_interview_trials(
            sigma, trials, weights, student_weight_data, course_graph_data,
            fixed_matches)
        print(f"Starting simulation {simulations} with sigma {sigma}")
        for i in range(1, 10):
            trial_matches = run_interview_trials(
                sigma, trials, weights, student_weight_data, course_graph_data,
                fixed_matches)
            sim_matches += trial_matches
            sim_trials += trials
            percent_decimals = trial_matches / trials
            if check_caps_interview_simulations(
                    course_data, sim_matches, i, sim_trials, percent_decimals):
                break
            trials *= 2
        simulation_percentages = sim_matches / sim_trials
        print(simulation_percentages.sum(axis=1))
        cumulative_percentages += simulation_percentages
        simulations += 1

    weighted_percentages = parse_interview_simulations(
        course_data, student_data, cumulative_percentages * 100.0 / simulations)
    write_interview_simulation(output_path, weighted_percentages)


def parse_interview_simulations(course_data: pd.DataFrame,
                                student_data: pd.DataFrame,
                                percentages: np.ndarray) -> Dict[
    str, List[Tuple[str, str, float]]]:
    readable_matches = {}
    courses = len(course_data.index)
    students = len(student_data.index)
    for ci in range(courses):
        sorted_matches = []
        for si in range(students):
            if percentages[si, ci] > 2.0:
                netid = student_data.index[si]
                name = student_data.iloc[si]['Name']
                percent = percentages[si, ci]
                sorted_matches.append((netid, name, percent))
        sorted_matches.sort(key=lambda x: x[2], reverse=True)
        readable_matches[course_data.index[ci]] = sorted_matches
    return readable_matches


def write_interview_simulation(output_path: str, readable_matches: Dict[
    str, List[Tuple[str, str, float]]]):
    with open(output_path + 'interview_simulations.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerow(['Course', 'NetID', 'Name', 'Percent Chance'])
        flattened = []
        for course, details in readable_matches.items():
            for netid, name, percentage in details:
                flattened.append([course, netid, name, percentage])
        writer.writerows(sorted(flattened, key=lambda x: x[0]))


def make_manual_adjustments(student_data: pd.DataFrame,
                            course_data: pd.DataFrame, weights: np.ndarray,
                            path_adjusted: str):
    if not os.path.isfile(path_adjusted):
        return
    adjusted_matches = pd.read_csv(
        path_adjusted, dtype={'NetID': str, 'Course': str, 'Weight': float})
    for _, row in adjusted_matches.iterrows():
        si = student_data.index.get_loc(row['NetID'])
        if row['Course'] not in course_data.index:
            continue
        ci = course_data.index.get_loc(row['Course'])
        weights[si, ci] += row['Weight']


def read_student_and_course_data(path: str, student_data: str,
                                 course_data: str) -> Tuple[
    pd.DataFrame, pd.DataFrame]:
    student_data = read_student_data(path + student_data)
    course_data = read_course_data(path + course_data)
    for student in student_data.index:
        if student not in course_data.columns.values:
            course_data[student] = np.nan
    for course in course_data.index:
        if course not in student_data.columns.values:
            student_data[course] = np.nan
    return student_data, course_data


def get_fixed_matches(course_data: pd.DataFrame, path: str,
                      student_data: pd.DataFrame) -> pd.DataFrame:
    if os.path.isfile(path):
        fixed_matches = pd.read_csv(path, dtype=str)
        fixed_matches['Student index'] = [student_data.index.get_loc(student)
                                          for student in fixed_matches['NetID']]
        fixed_matches['Course index'] = [course_data.index.get_loc(course) for
                                         course in fixed_matches['Course']]
    else:
        fixed_matches = pd.DataFrame(
            columns=['NetID', 'Course', 'Student index', 'Course index'])
    return fixed_matches


def make_adjustments_from_previous(course_data: pd.DataFrame,
                                   path_to_previous: str,
                                   student_data: pd.DataFrame,
                                   weights: np.ndarray) -> pd.DataFrame:
    if not os.path.isfile(path_to_previous):
        return pd.DataFrame(columns=['NetID', 'Course'])
    previous_matches = pd.read_csv(
        path_to_previous, dtype={'NetID': str, 'Course': str})
    for _, row in previous_matches.iterrows():
        if row['NetID'] not in student_data.index or row[
            'Course'] not in course_data.index:
            continue
        si = student_data.index.get_loc(row['NetID'])
        ci = course_data.index.get_loc(row['Course'])
        weights[si, ci] += params.PREVIOUS_MATCHING_BOOST
    return previous_matches


def run_alternate_matchings(alternates: int, course_data: pd.DataFrame,
                            fixed_matches: pd.DataFrame,
                            best_matches: List[Tuple[int, int]], path: str,
                            student_data: pd.DataFrame, weights: np.ndarray) -> \
        List[float]:
    last_matches = best_matches
    cumulative = 0.0
    alt_weights = []
    for i in range(alternates):
        last_matches, cumulative, alt_weight = find_alternate_matching(
            f'{path}alternate{i + 1}.csv', student_data, course_data, weights,
            fixed_matches, best_matches, last_matches, cumulative)
        alt_weights.append(alt_weight)
    return alt_weights


def validate_path_args(path: str, output: str) -> str:
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
        '--previous', metavar='PREVIOUS MATCHING',
        default='inputs/previous.csv',
        help='csv file with previous matching algorithm execution output')
    parser.add_argument(
        '--output', metavar='MATCHING OUTPUT', default='outputs/',
        help='location to write matching output')
    parser.add_argument(
        '--alternates', metavar='ALTERNATE MATCHES', type=int, default=3,
        help='number of alternate matches to solve for')
    parser.add_argument(
        '--run_interviews', default=False, action='store_true',
        help='run the interviews simulation')
    args = parser.parse_args()
    args.path = validate_path_args(args.path, args.output)

    run_matching(**vars(args))
