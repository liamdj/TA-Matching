import copy
import datetime
import random
import string
from typing import List, Tuple

import g_sheet_consts as gs_const
import write_to_google_sheets as write_gs

COURSES = ['COS 126', 'COS 217', 'COS 226', 'COS 240', 'COS 302', 'COS 316',
           'COS 320', 'COS 324', 'COS 333', 'COS 418', 'COS 426', 'COS 432',
           'COS 445', 'COS 475', 'COS 484', 'COS 488', 'COS 495', 'COS 511',
           'COS 518', 'COS 522', 'COS 585', 'EGR 154']


def _weighted_rand(first, last, weights):
    return random.choices(
        [i for i in range(first, last + 1)], weights=weights, k=1)[0]


def gen_princeton_email():
    return ''.join(
        [random.choice(string.ascii_lowercase + string.digits) for _ in
         range(random.randint(7, 12))]) + '@princeton.edu'


def gen_name():
    def gen_first_letter():
        return random.choice(string.ascii_uppercase)

    def gen_next_letters():
        return random.choice(string.ascii_lowercase)

    name = []
    name.append(gen_first_letter())
    for _ in range(random.randint(1, 9)):
        name.append(gen_next_letters())
    return ''.join(name)


def gen_nickname():
    return gen_name() if _weighted_rand(0, 1, (9, 1)) else ""


def gen_advisor2():
    return gen_name() if _weighted_rand(0, 1, (14, 1)) else ""


def gen_track_year():
    if _weighted_rand(0, 1, (2, 5)):
        return ("MSE", random.randint(1, 2))
    return ("PHD", _weighted_rand(1, 4, (5, 5, 1, 1)))


def gen_bank_join_score(track: str):
    option = _weighted_rand(0, 2, (5, 1, 1))
    bank, join = "", ""
    if 'PHD' == track:
        if option == 1:
            bank = str(random.randint(2, 10) / 2.0)
        elif option == 2:
            join = str(random.randint(2, 10) / 2.0)
    return bank, join


def gen_half_score():
    return str(random.randint(2, 10) / 2.0) if _weighted_rand(
        0, 1, (15, 1)) else ""


def gen_course_name():
    return COURSES[random.randint(0, len(COURSES) - 1)]


def parse_course_num(course_name):
    return course_name[4:]


def gen_course_num():
    return parse_course_num(gen_course_name())


def gen_match_list():
    return ', '.join(
        [gen_course_name() for _ in range(
            random.choices([1, 2, 3, 4], weights=(4, 3, 2, 1), k=1)[0])])


def gen_course_assignment():
    return "" if _weighted_rand(0, 1, (1, 11)) else gen_course_num()


def gen_previously_taught_courses(at_princeton):
    last = 6
    courses = []
    while random.randint(0, last) >= 2:
        c = gen_course_name()
        term = random.choices(["Fall", "Spring"])[0] + str(
            random.randint(17, 21))
        if at_princeton:
            courses.append(f"{c} ({term})")
        else:
            courses.append(f"{c} ({term}, {gen_name()})")
        last -= 1
    return "; ".join(courses)


def gen_timestamp():
    return datetime.datetime.now().strftime('%m/%d/%Y %H:%M:%S')


def generate_faculty_sheet(entries=50):
    titles = [['Key', 'NetID', 'Last', 'First', 'Department', 'Courses']]
    matrix = []
    courses = copy.deepcopy(COURSES)
    random.shuffle(courses)
    for _ in range(entries):
        last_name = gen_name()
        course = []
        for _ in range(random.choices([0, 1, 2], weights=(25, 20, 1), k=1)[0]):
            if len(courses):
                course.append(parse_course_num(courses.pop()))
        matrix.append(
            [last_name, gen_name().lower(), last_name, gen_name(), 'COS',
             ','.join(course)])

    matrix = sorted(matrix, key=lambda x: x[0])
    return titles + matrix


def generate_universal_student_matrix(entries, advisors_matrix):
    matrix = [("First Name", "Last Name", "NetID", "Comma Separated Advisors")]
    advisors_matrix = advisors_matrix[1:]
    for _ in range(entries):
        advisor1 = random.choice(advisors_matrix)[2]  # just last name
        advisor2 = random.choice(advisors_matrix)[2]
        advisors = advisor1 if _weighted_rand(
            0, 1, (1, 14)) else f"{advisor1}, {advisor2}"
        matrix.append((gen_name(), gen_name(), gen_name().lower(), advisors))
    return matrix


def generate_student_info_sheet(entries=5):
    titles = [["Last", "First", "Nickname", "NetID", "Form", "Track", "Year",
               "Advisor", "Advisor2", "Bank", "Join", "Half", "Course", "Notes",
               "Special Notes"]]
    matrix = []
    for _ in range(entries):
        track, year = gen_track_year()
        line = [gen_name(), gen_name(), gen_nickname(), gen_name().lower(),
                "yes", track, year, gen_name(), gen_advisor2(),
                *gen_bank_join_score(track), gen_half_score(),
                gen_course_assignment(), "", ""]
        matrix.append(line)

    matrix = sorted(matrix, key=lambda x: x[0])
    return titles + matrix


def generate_student_info_sheet_from_matrix(student_info):
    # student = ["First Name", "Last Name", "NetID", "Comma Separated Advisors"]
    titles = [["Last", "First", "Nickname", "NetID", "Form", "Track", "Year",
               "Advisor", "Advisor2", "Bank", "Join", "Half", "Course", "Notes",
               "Special Notes"]]
    matrix = []
    for student in student_info[1:]:
        track, year = gen_track_year()
        advisors = student[3].split(", ")
        advisor1 = advisors[0]
        advisor2 = "" if len(advisors) == 1 else advisors[1]
        line = [student[0], student[1], gen_nickname(), student[2], "yes",
                track, year, advisor1, advisor2, *gen_bank_join_score(track),
                gen_half_score(), gen_course_assignment(), "", ""]
        matrix.append(line)
    matrix = sorted(matrix, key=lambda x: x[0])
    return titles + matrix


def generate_student_preferences(student_info_sheet, advisors_matrix):
    """
    student has format ["Last", "First", "Nickname", "NetID", "Form", "Track", "Year", "Advisor", "Advisor2", "Bank", "Join", "Half", "Course", "Notes", "Special Notes"]

    advisor has format ["Key", "NetID", "Last", "First", "Department", "Courses"]
    """
    PREVIOUS_NON_PRINCETON_COURSES_TITLE = """What courses have you previously been a (graduate or undergraduate) TA for, at any other university? If none, leave blank. Include equivalent courses using the Princeton number equivalent. Please use the following format: COS NUM1 (TERM1, UNIVERSITY1); COS NUM2 (TERM2, UNIVERSITY2). An example answer might look like: COS 126 (Spring18, Cornell); COS 226 (Fall20, Cornell). """
    PREVIOUS_PRINCETON_COURSES_TITLE = """What courses at Princeton have you previously been a (graduate or undergraduate) TA for? If none, leave blank. Please use the following format: COS NUM1 (TERM1); COS NUM2 (TERM2). An example answer might look like: COS 126 (Spring18); COS 226 (Fall20); COS 226 (Spring21). """
    OTHER_COMMENTS_TITLE = """Any other comments that might be helpful for someone who is considering you for a TA position? (This will be shared with all instructors.) It is fine to leave this blank."""
    OTHER_PRIVATE_COMMENTS_TITLE = """Any other private comments for us as we optimize over TA assignments? (This will be shared with Adam Finkelstein and Matt Weinberg and possibly one or two others involved in TA assignments.) It is fine to leave this blank."""
    matrix = [["Timestamp", "Email", "Full Name", "Advisor",
               PREVIOUS_PRINCETON_COURSES_TITLE,
               PREVIOUS_NON_PRINCETON_COURSES_TITLE, "Favorite Match",
               "Good Match", "OK Match", OTHER_COMMENTS_TITLE,
               OTHER_PRIVATE_COMMENTS_TITLE]]
    advisors_matrix = advisors_matrix[1:]
    advisors_key_to_full = {}
    for advisor in advisors_matrix:
        advisors_key_to_full[advisor[0]] = f"{advisor[3]} {advisor[2]}"

    for student in student_info_sheet[1:]:
        advisors = []
        for advisor in [student[7], student[8]]:
            if advisor:
                advisors.append(advisors_key_to_full[advisor])

        line = [gen_timestamp(), f"{student[3]}@princeton.edu",
                f"{student[1]} {student[0]}", "; ".join(advisors),
                gen_previously_taught_courses(True),
                gen_previously_taught_courses(False), gen_match_list(),
                gen_match_list(), gen_match_list(), "", "", ""]
        matrix.append(line)
    return matrix


def generate_fac_preferences(student_preferences, advisors_matrix) -> List[
    List[str]]:
    def __gen_matches_list(possible_students, is_best=True):
        if is_best:
            n = min(
                len(possible_students), random.choices(
                    [0, 1, 3, 6, 10, 14], weights=[1, 1, 3, 5, 6, 3], k=1)[0])
        else:
            n = min(
                len(possible_students),
                random.choices([0, 1, 3, 5], weights=[6, 5, 3, 1], k=1)[0])
        matches = random.choices(possible_students, k=n)
        return ', '.join(matches)

    matrix = [["Timestamp", "Email Address", "Which course?", "Best Match(es)",
               "Matches to Avoid"]]
    courses = []
    courses_to_students = {}
    for course in COURSES:
        if _weighted_rand(0, 1, (1, 3)):
            courses.append(course)
        courses_to_students[course] = []

    for student in student_preferences[1:]:
        for i in range(6, 9):
            for course in student[i].split(", "):
                courses_to_students[course].append(
                    f"{student[2]} <{student[1]}>")

    teaching_advisors_netids = []
    for advisor in advisors_matrix[1:]:
        if advisor[5]:
            teaching_advisors_netids.append(advisor[1])

    for course in courses:
        email = random.choice(teaching_advisors_netids) + '@princeton.edu'
        possible_students = courses_to_students[course]
        line = [gen_timestamp(), email, course,
                __gen_matches_list(possible_students),
                __gen_matches_list(possible_students, False)]
        matrix.append(line)
    return matrix


def generate_courses_matrix(faculty_matrix):
    """ faculty_matrix has format ['Key', 'NetID', 'Last', 'First', 'Department', 'Courses'] """
    titles = [
        ["Course", "Omit", "TAs", "Weight", "Instructor", "Title", "Notes"]]
    faculty = [fac[0] for fac in faculty_matrix]
    random.shuffle(faculty)

    matrix = []
    for course in COURSES:
        num = course.replace(' ', '')
        if 'COS' in course:
            num = num[3:]

        instructor = []
        for _ in range(_weighted_rand(1, 2, (25, 1))):
            if len(faculty):
                instructor.append(faculty.pop())
        instructor = ','.join(instructor)

        omit = random.choices(['x', ''], weights=(1, 25), k=1)[0]
        title = f"{gen_name()} {gen_name()} {gen_name()}"
        TAs = _weighted_rand(1, 12, range(12, 0, -1))
        weight = _weighted_rand(0, 2, (25, 1, 1))

        line = [num, omit, TAs, weight if weight else '', instructor, title, '']
        matrix.append(line)

    matrix = sorted(matrix, key=lambda x: x[0])
    return titles + matrix


def gen_full_input_matrices(student_entries=5, faculty_entries=50) -> Tuple[
    List[List[str]], List[List[str]], List[List[str]], List[List[str]], List[
        List[str]]]:
    faculty = generate_faculty_sheet(faculty_entries)
    courses = generate_courses_matrix(faculty)
    student_matrix = generate_universal_student_matrix(
        student_entries, faculty)
    student_info = generate_student_info_sheet_from_matrix(student_matrix)
    student_prefs, fac_prefs = gen_preferences_from_student_info_and_advisors(
        student_info, faculty)
    return student_info, faculty, courses, student_prefs, fac_prefs


def gen_preferences_from_student_info_and_advisors(student_info, advisors):
    student_preferences = generate_student_preferences(student_info, advisors)
    fac_preferences = generate_fac_preferences(student_preferences, advisors)
    return student_preferences, fac_preferences


def generate_and_write_all_input_sheets(student_entries=5,
                                        faculty_entries=50) -> Tuple[
    str, str, str]:
    student_info, faculty, courses, student_prefs, fac_prefs = gen_full_input_matrices(
        student_entries, faculty_entries)
    student_info_sheet_name = "TA Matching 22: TA Planning (Generated Randomly)"
    planning_sheet = write_gs.write_matrix_to_sheet(
        student_info, student_info_sheet_name, "Students")
    write_gs.write_matrix_to_new_tab(faculty, planning_sheet, "Faculty")
    write_gs.write_matrix_to_new_tab(courses, planning_sheet, "Courses")
    s_pref_sheet_id, i_pref_sheet_id = write_student_and_fac_preferences(
        student_prefs, fac_prefs)
    return planning_sheet.id, s_pref_sheet_id, i_pref_sheet_id


def write_student_and_fac_preferences(student_preferences, fac_preferences) -> \
        Tuple[str, str]:
    student_prefs_sheet = write_gs.write_matrix_to_sheet(
        student_preferences, gs_const.GENERATED_TA_PREFS_SHEET_TITLE,
        gs_const.PREFERENCES_INPUT_TAB_TITLE)
    fac_prefs_sheet = write_gs.write_matrix_to_sheet(
        fac_preferences, gs_const.GENERATED_INSTRUCTORS_PREFS_SHEET_TITLE,
        gs_const.PREFERENCES_INPUT_TAB_TITLE, wrap=True)
    return student_prefs_sheet.id, fac_prefs_sheet.id


def generate_and_write_preferences_from_student_info_and_advisors(
        student_info_sheet_id):
    sheet = write_gs.get_sheet_by_id(student_info_sheet_id)
    student_info = sheet.worksheet("Students").get_all_values()
    advisors = sheet.worksheet("Faculty").get_all_values()
    student_prefs, fac_prefs = gen_preferences_from_student_info_and_advisors(
        student_info, advisors)
    write_student_and_fac_preferences(student_prefs, fac_prefs)
