import csv
from typing import Dict, Tuple, List

import g_sheet_consts
import write_to_google_sheets as write_gs


def compare_matchings(old_matches: Dict[str, str],
                      new_matches: Dict[str, str]) -> Tuple[
    Dict[str, Tuple[str, str]], Dict[str, Tuple[List[str], List[str]]]]:
    """
    Take in the Name -> course matchings and output (student_changes, course_changes),
    where student_changes = Name -> (old_course, new_course),
    and course_changes = Course -> (old_students, new_students)
    """
    student_changes = {}
    course_changes = {}
    for student, course in new_matches.items():
        if old_matches.get(student):
            if old_matches[student] != course:
                student_changes[student] = (old_matches[student], course)
            del old_matches[student]
        else:
            student_changes[student] = (None, course)

    for student, old_course in old_matches.items():
        student_changes[student] = (old_course, None)

    for student, (old_course, new_course) in student_changes.items():
        if old_course in course_changes.keys():
            old_students, _ = course_changes[old_course]
            old_students.append(student)
        elif old_course is not None and old_course != 'unassigned':
            course_changes[old_course] = ([student], [])
        if new_course in course_changes.keys():
            _, new_students = course_changes[new_course]
            new_students.append(student)
        elif new_course is not None and new_course != 'unassigned':
            course_changes[new_course] = ([], [student])

    return student_changes, course_changes


def read_matching_from_csv(csv_filename: str) -> Dict[str, str]:
    with open(csv_filename) as f:
        reader = csv.DictReader(f)
        return read_matching_from_matrix(reader)


def read_matching_from_matrix(matrix) -> Dict[str, str]:
    name_to_course = {}
    for row in matrix:
        name_to_course[row['Name']] = row['Course']
    return name_to_course


def compare_matching_csvs(old_matching_csv_filename: str,
                          new_matching_csv_filename: str) -> Tuple[
    Dict[str, Tuple[str, str]], Dict[str, Tuple[List[str], List[str]]]]:
    old_matching = read_matching_from_csv(old_matching_csv_filename)
    new_matching = read_matching_from_csv(new_matching_csv_filename)
    return compare_matchings(old_matching, new_matching)


def compare_matching_worksheet_with_csv(old_matching_worksheet,
                                        new_matching_csv_filename: str) -> \
        Tuple[
            Dict[str, Tuple[str, str]], Dict[str, Tuple[List[str], List[str]]]]:
    old_matching = read_matching_from_matrix(
        old_matching_worksheet.get_all_records())
    new_matching = read_matching_from_csv(new_matching_csv_filename)
    return compare_matchings(old_matching, new_matching)


def compare_matching_worksheets(old_matching_worksheet_title: str,
                                new_matching_worksheet_title: str) -> Tuple[
    Dict[str, Tuple[str, str]], Dict[str, Tuple[List[str], List[str]]]]:
    sheet = write_gs.get_sheet(g_sheet_consts.MATCHING_OUTPUT_SHEET_TITLE)
    old_matching_worksheet = write_gs.get_worksheet_from_sheet(
        sheet, old_matching_worksheet_title)
    old_matching = read_matching_from_matrix(
        old_matching_worksheet.get_all_records())
    new_matching_worksheet = write_gs.get_worksheet_from_sheet(
        sheet, new_matching_worksheet_title)
    new_matching = read_matching_from_matrix(
        new_matching_worksheet.get_all_records())
    return compare_matchings(old_matching, new_matching)


def write_comparison_to_worksheet(student_changes: Dict[str, Tuple[str, str]],
                                  course_changes: Dict[
                                      str, Tuple[List[str], List[str]]],
                                  new_worksheet_title: str):
    sheet = write_gs.get_sheet(g_sheet_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE)
    worksheets = {ws.title: ws for ws in sheet.worksheets()}
    prefix = f'https://docs.google.com/spreadsheets/d/{sheet.id}#gid='
    if new_worksheet_title + '(C)' in worksheets:
        student_ws_id = worksheets[new_worksheet_title + '(C)'].id
        print(f"Course comparison (already existed): {prefix}{student_ws_id}")
    else:
        course_changes_ws = write_gs.create_and_write_full_worksheet(
            course_changes_to_matrix(course_changes), sheet,
            new_worksheet_title + '(C)')
        print(f"Course comparison: {prefix}{course_changes_ws.id}")

    if new_worksheet_title + '(S)' in worksheets:
        student_ws_id = worksheets[new_worksheet_title + '(S)'].id
        print(f"Student comparison (already existed): {prefix}{student_ws_id}")
    else:
        student_changes_ws = write_gs.create_and_write_full_worksheet(
            student_changes_to_matrix(student_changes), sheet,
            new_worksheet_title + '(S)')
        print(f"Student comparison: {prefix}{student_changes_ws.id}")


def student_changes_to_matrix(student_changes: Dict[str, Tuple[str, str]]) -> \
        List[List[str]]:
    matrix = [['Name', 'Old Course', 'New Course']]
    for name, (old_course, new_course) in student_changes.items():
        old_course = 'added' if old_course is None else old_course
        new_course = 'removed' if new_course is None else new_course
        matrix.append([name, old_course, new_course])
    return matrix


def course_changes_to_matrix(
        course_changes: Dict[str, Tuple[List[str], List[str]]]) -> List[
    List[str]]:
    matrix = [['Course', 'Removed TAs', 'Added TAs']]
    for course, (old_students, new_students) in course_changes.items():
        matrix.append(
            [course, '; '.join(old_students), '; '.join(new_students)])
    return matrix


def write_matchings_changes(student_changes: Dict[str, Tuple[str, str]],
                            course_changes: Dict[
                                str, Tuple[List[str], List[str]]],
                            output_dir='data'):
    with open(output_dir + '/outputs/matchings_students_diff.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerows(student_changes_to_matrix(student_changes))

    with open(output_dir + '/outputs/matchings_courses_diff.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerows(course_changes_to_matrix(course_changes))

    return len(student_changes)
