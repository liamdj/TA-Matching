import write_to_google_sheets as write_gs
import datetime
import random
import string
import copy

COURSES = ['COS 126', 'COS 217', 'COS 226', 'COS 240', 'COS 302', 'COS 316', 'COS 320', 'COS 324', 'COS 333',  'COS 418',
           'COS 426', 'COS 432', 'COS 445', 'COS 475', 'COS 484', 'COS 488', 'COS 495', 'COS 511', 'COS 518', 'COS 522', 'COS 585', 'EGR 154']


def _weighted_rand(first, last, weights):
    return random.choices([i for i in range(first, last+1)], weights=weights, k=1)[0]


def gen_princeton_email():
    return ''.join([random.choice(string.ascii_lowercase + string.digits) for _ in range(random.randint(7, 12))]) + '@princeton.edu'


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


def gen_bank_join_score():
    return str(random.randint(2, 10)/2.0) if _weighted_rand(0, 1, (4, 1)) else ""


def gen_course_name():
    return COURSES[random.randint(0, len(COURSES) - 1)]


def parse_course_num(course_name):
    return course_name[4:]


def gen_course_num():
    return parse_course_num(gen_course_name())


def gen_match_list():
    return ', '.join([gen_course_name() for _ in range(random.choices([1, 2, 3, 4], weights=(4, 3, 2, 1), k=1)[0])])


def gen_course_assignment():
    return "" if _weighted_rand(0, 1, (1, 11)) else gen_course_num()


def gen_previously_taught_courses():
    last = 6
    courses = []
    while random.randint(0, last) >= 2:
        c = gen_course_name()
        term = random.choices(["Fall", "Spring"])[0] + \
            str(random.randint(17, 21))
        uni = random.choices([gen_name(), "Princeton"])[0]
        courses.append(f"{c} ({term}, {uni})")
        last -= 1
    return "; ".join(courses)


def gen_timestamp():
    return datetime.datetime.now().strftime('%m/%d/%Y %H:%M:%S')


def generate_advisors_sheet(entries=50):
    out = [['Key', 'NetID', 'Last', 'First', 'Department', 'Courses']]
    matrix = []
    courses = copy.deepcopy(COURSES)
    random.shuffle(courses)
    for _ in range(entries):
        last_name = gen_name()
        course = []
        for _ in range(random.choices([0, 1, 2], weights=(25, 20, 1), k=1)[0]):
            if len(courses):
                course.append(parse_course_num(courses.pop()))
        matrix.append([last_name, gen_name().lower(),
                      last_name, gen_name(), 'COS', ','.join(course)])
    matrix = sorted(matrix, key=lambda x: x[0])
    return out + matrix


def generate_universal_student_matrix(entries, advisors_matrix):
    matrix = [("First Name", "Last Name", "NetId", "Comma Separated Advisors")]
    advisors_matrix = advisors_matrix[1:]
    for _ in range(entries):
        advisor1 = random.choice(advisors_matrix)[2]  # just last name
        advisor2 = random.choice(advisors_matrix)[2]
        advisors = advisor1 if _weighted_rand(
            0, 1, (1, 14)) else f"{advisor1}, {advisor2}"
        matrix.append((gen_name(), gen_name(), gen_name().lower(), advisors))
    return matrix


def generate_student_info_sheet(entries=5):
    out = [["Last", "First", "Nickname", "NetId", "Form", "Track",
            "Year", "Advisor", "Advisor2", "Bank", "Join", "Course", "Notes"]]
    matrix = []
    for _ in range(entries):
        track, year = gen_track_year()
        line = [gen_name(), gen_name(), gen_nickname(),
                gen_name().lower(), "yes", track, year, gen_name(), gen_advisor2(), gen_bank_join_score(), gen_bank_join_score(), gen_course_assignment(), ""]
        matrix.append(line)
    matrix = sorted(matrix, key=lambda x: x[0])
    return out + matrix


def generate_student_info_sheet_from_matrix(student_info):
    # student = ["First Name", "Last Name", "NetId", "Comma Separated Advisors"]
    matrix = []
    for student in student_info[1:]:
        track, year = gen_track_year()
        advisors = student[3].split(", ")
        advisor1 = advisors[0]
        advisor2 = "" if len(advisors) == 1 else advisors[1]
        line = [student[0], student[1], gen_nickname(),
                student[2], "yes", track, year, advisor1, advisor2, gen_bank_join_score(), gen_bank_join_score(), gen_course_assignment(), ""]
        matrix.append(line)
    out = [["Last", "First", "Nickname", "NetId", "Form", "Track",
            "Year", "Advisor", "Advisor2", "Bank", "Join", "Course", "Notes"]]
    matrix = sorted(matrix, key=lambda x: x[0])
    return out + matrix


def generate_student_preferences(student_info_sheet, advisors_matrix):
    """
    student has format ["Last", "First", "Nickname", "NetId", "Form", "Track", "Year", "Advisor", "Advisor2", "Bank", "Join", "Course", "Notes"]

    advisor has format ["Key", "NetID", "Last", "First", "Department", "Courses"]
    """
    PREVIOUS_COURSES_TITLE = """What courses have you previously been a (graduate or undergraduate) TA for? If none, leave blank. Include equivalent courses at your undergraduate institution, using the Princeton number equivalent. *Please use the following format*: COS NUM1 (TERM1, UNIVERSITY1); COS NUM2 (TERM2, UNIVERSITY2). An example answer might look like: COS 126 (Spring18, Cornell); COS 226 (Fall20, Princeton); COS 226 (Spring21, Princeton)."""
    OTHER_COMMENTS_TITLE = """Any other comments that might be helpful for someone who is considering you for a TA position? (This will be shared with all instructors.) It is fine to leave this blank."""
    OTHER_PRIVATE_COMMENTS_TITLE = """Any other private comments for us as we optimize over TA assignments? (This will be shared with Adam Finkelstein and Matt Weinberg and possibly one or two others involved in TA assignments.) It is fine to leave this blank."""
    matrix = [["Timestamp", "Email", "Full Name", "Advisor",
               PREVIOUS_COURSES_TITLE, "Favorite Match", "Good Match", "OK Match", OTHER_COMMENTS_TITLE, OTHER_PRIVATE_COMMENTS_TITLE]]
    advisors_matrix = advisors_matrix[1:]
    advisors_key_to_full = {}
    for advisor in advisors_matrix:
        advisors_key_to_full[advisor[0]] = f"{advisor[3]} {advisor[2]}"

    for student in student_info_sheet[1:]:
        advisors = []
        for advisor in [student[7], student[8]]:
            if advisor:
                advisors.append(advisors_key_to_full[advisor])

        line = [gen_timestamp(), f"{student[3]}@princeton.edu", f"{student[1]} {student[0]}", "; ".join(advisors),
                gen_previously_taught_courses(), gen_match_list(), gen_match_list(), gen_match_list(), "", ""]
        matrix.append(line)
    return matrix


def generate_fac_preferences(student_preferences, advisors_matrix):
    def __gen_matches_list(possible_students, is_best=True):
        if is_best:
            n = min(len(possible_students), random.choices(
                [0, 1, 3, 6, 10, 14], weights=[1, 1, 3, 5, 6, 3], k=1)[0])
        else:
            n = min(len(possible_students), random.choices(
                [0, 1, 3, 5], weights=[6, 5, 3, 1], k=1)[0])
        matches = random.choices(possible_students, k=n)
        return ', '.join(matches)

    matrix = [["Timestamp", "Email Address", "Which course?",
               "Best Match(es)", "Matches to Avoid"]]
    courses = []
    courses_to_students = {}
    for course in COURSES:
        if _weighted_rand(0, 1, (1, 3)):
            courses.append(course)
        courses_to_students[course] = []

    for student in student_preferences[1:]:
        for i in range(5, 8):
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
                __gen_matches_list(possible_students), __gen_matches_list(possible_students)]
        matrix.append(line)
    return matrix


def gen_full_input_sheets(student_entries=5, advisors_entries=50):
    advisors_sheet = generate_advisors_sheet(advisors_entries)
    student_matrix = generate_universal_student_matrix(
        student_entries, advisors_sheet)
    student_info_sheet = generate_student_info_sheet_from_matrix(
        student_matrix)
    student_prefs, fac_prefs = gen_preferences_from_student_info_and_advisors(
        student_info_sheet, advisors_sheet)
    return student_info_sheet, student_prefs, fac_prefs, advisors_sheet


def gen_preferences_from_student_info_and_advisors(student_info, advisors):
    student_preferences = generate_student_preferences(
        student_info, advisors)
    fac_preferences = generate_fac_preferences(
        student_preferences, advisors)
    return student_preferences, fac_preferences


def generate_and_write_all_input_sheets(student_entries=5, advisors_entries=50):
    student_info_sheet, student_preferences, fac_preferences, advisors_sheet = gen_full_input_sheets(
        student_entries, advisors_entries)
    student_info_sheet_name = "TA Matching 22: Student Info (Generated Randomly)"
    write_gs.write_matrix_to_sheet(
        student_info_sheet, student_info_sheet_name, "Students")
    write_gs.write_matrix_to_new_tab(
        advisors_sheet, student_info_sheet_name, "Advisors")
    write_student_and_fac_preferences(student_preferences, fac_preferences)


def write_student_and_fac_preferences(student_preferences, fac_preferences):
    student_prefs_sheet_id = write_gs.write_matrix_to_sheet(
        student_preferences, "TA Matching 22: Student Preferences (Generated Randomly)", "Form Responses 1")
    fac_prefs_sheet_id = write_gs.write_matrix_to_sheet(
        fac_preferences, "TA Matching 22: Instructor Preferences (Generated Randomly)", "Form Responses 1", wrap=True)
    return student_prefs_sheet_id, fac_prefs_sheet_id


def generate_and_write_preferences_from_student_info_and_advisors(student_info_sheet_id):
    sheet = write_gs.get_sheet_by_id(student_info_sheet_id)
    student_info = sheet.worksheet("Students").get_all_values()
    advisors = sheet.worksheet("Advisors").get_all_values()
    student_prefs, fac_prefs = gen_preferences_from_student_info_and_advisors(
        student_info, advisors)
    write_student_and_fac_preferences(student_prefs, fac_prefs)