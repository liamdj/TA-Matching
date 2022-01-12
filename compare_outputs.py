import csv
from typing import Dict, List, Tuple
from pprint import pprint


def compare_matchings(old_matches: Dict[str, str],
                      new_matches: Dict[str, str]) -> Tuple[
    Dict[str, Tuple[str, str]], Dict[str, Tuple[List[str], List[str]]]]:
    """
    Take in the Netid -> course matchings and output (student_changes, course_changes),
    where student_changes = Netid -> (old_course, new_course),
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
    netid_to_course = {}
    for row in matrix:
        netid_to_course[row['Name']] = row['Course']
    return netid_to_course


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


def write_matchings_changes(student_changes: Dict[str, Tuple[str, str]],
                            course_changes: Dict[
                                str, Tuple[List[str], List[str]]],
                            output_dir='data'):
    with open(output_dir + '/outputs/matchings_students_diff.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerow(['Netid', 'Old Course', 'New Course'])
        for netid, (old_course, new_course) in student_changes.items():
            old_course = 'added' if old_course is None else old_course
            new_course = 'removed' if new_course is None else new_course
            writer.writerow([netid, old_course, new_course])

    with open(output_dir + '/outputs/matchings_courses_diff.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerow(['Course', 'Removed TAs', 'Added TAs'])
        for course, (old_students, new_students) in course_changes.items():
            writer.writerow(
                [course, '; '.join(old_students), '; '.join(new_students)])

    return len(student_changes)
