import os
import re
import datetime
import csv
from typing import Union, Any, TypedDict, Optional

import gspread

import g_sheet_consts as gs_consts


class StudentType(TypedDict, total=False):
    NetID: str
    Name: str
    Advisor: str
    Previous: str
    Favorite: str
    Good: str
    OK: str
    Bank: Union[str, float]
    Join: Union[str, float]
    Notes: str


class StudentInfoValue(TypedDict):
    Last: str
    First: str
    Nickname: str
    NetID: str
    Form: str
    Dept: str
    Track: str
    Year: int
    Advisor: str
    Advisor2: str
    Bank: Union[str, float]
    Join: Union[str, float]
    Half: Union[str, float]
    Course: str
    Notes: str
    Shorthand: str


class CourseValue(TypedDict):
    Course: Union[str, int]
    TAs: int
    Weight: Union[str, float]
    Instructor: str
    Title: str
    Notes: str


class FacultyPref(TypedDict):
    Timestamp: str
    Email: str
    Course: str
    Favorite: str
    Veto: str


StudentsType = dict[str, StudentType]
StudentInfoType = list[StudentInfoValue]
CoursesType = dict[str, CourseValue]
FacultyPrefsType = dict[str, FacultyPref]
AdjustedType = dict[str, set[str]]
AssignedType = dict[str, str]
NamesType = dict[str, str]
NotesType = dict[str, str]
YearsType = dict[str, str]


def get_rows_with_tab_title(sheet_id: str, tab_title: str) -> Union[
    list, list[dict]]:
    gc = gspread.service_account(filename='./credentials.json')

    sheet = gc.open_by_key(sheet_id).worksheet(tab_title)
    all_rows = sheet.get_all_records()
    return all_rows


def parse_pre_assign(student_info: StudentInfoType) -> AssignedType:
    fixed = {}
    for row in student_info:
        if _sanitize(row["Course"]):
            fixed[row["NetID"]] = add_COS(row["Course"])
    return fixed


def _sanitize(raw: Any) -> Any:
    return raw if type(raw) != str else raw.strip()


def parse_years(student_info: StudentInfoType) -> YearsType:
    years = {}
    for student in student_info:
        netid = student["NetID"]
        track = student["Track"]
        year = student["Year"]
        years[netid] = track + str(year)
    return years


def add_COS(courses: str) -> str:
    """ Ignores COS """
    parts = str(courses).split(';')
    for i, part in enumerate(parts):
        # if course code from a non-COS discipline
        if re.match(r'[A-Z]', part):
            parts[i] = part.replace(' ', '')
        else:
            parts[i] = 'COS' + part
    courses = ';'.join(parts)
    return courses


def parse_student_preferences(
        student_prefs: list[dict[str, str]]) -> StudentsType:
    students = {}
    cols = {'Email': 'NetID', 'Name': 'Name', 'Advisor': 'Advisor',
            'What courses at Princeton': 'Previous', 'Favorite': 'Favorite',
            'Good': 'Good', 'OK Match': 'OK'}

    student_prefs = switch_keys_from_rows(student_prefs, cols, True)
    for row in student_prefs:
        fav = format_course_list(row['Favorite'])
        good = format_course_list(row['Good'])
        okay = format_course_list(row['OK'])
        fav = fav.split(';')
        good = set(good.split(';'))
        okay = set(okay.split(';'))
        for fav_c in fav:
            if fav_c in good:
                good.remove(fav_c)
            if fav_c in okay:
                okay.remove(fav_c)
        for good_c in good:
            if good_c in okay:
                okay.remove(good_c)

        row['Favorite'] = ';'.join(fav)
        row['Good'] = ';'.join(good)
        row['OK'] = ';'.join(okay)
        row['NetID'] = format_netid(row['NetID'])
        students[row['NetID']] = row
    return students


def parse_courses(course_info: list[dict]) -> CoursesType:
    courses = {}
    for row in course_info:
        courses[row['Course']] = row
    return courses


def switch_keys(dictionary: dict[str, Any],
                key_to_switch_dict: dict[str, str]) -> dict[str, Any]:
    new = {}
    for key, val in dictionary.items():
        if key in key_to_switch_dict:
            new[key_to_switch_dict[key]] = val
    return new


def switch_keys_from_rows(rows: list[dict], keys_to_switch_dict: dict[str, str],
                          is_paraphrased=False) -> list[dict]:
    """
    If `is_paraphrased`, then assume that the keys in `keys_to_switch_dict`
    are only key phrases in the full keys in `rows` (that will only be 
    replicated in one key), and so this function will translate from the 
    paraphrased keys to the full keys.
    """

    new_rows = []
    if is_paraphrased:
        full_keys_to_switch = {}  # from key term to full keys
        for key_term, val in keys_to_switch_dict.items():
            for full_key in rows[0].keys():
                if key_term in full_key:
                    full_keys_to_switch[full_key] = val
        keys_to_switch_dict = full_keys_to_switch

    for row in rows:
        new_rows.append(switch_keys(row, keys_to_switch_dict))
    return new_rows


def parse_faculty_prefs(
        faculty_prefs: list[dict[str, str]]) -> FacultyPrefsType:
    cols = {'Email': 'Email', 'Which course?': 'Course', 'Best': 'Favorite',
            'Avoid': 'Veto'}
    courses = {}
    faculty_prefs = switch_keys_from_rows(faculty_prefs, cols, True)
    for row in faculty_prefs:
        courses[row['Course']] = row
    return courses


def fix_advisors(students: StudentsType,
                 student_info: StudentInfoType) -> StudentsType:
    email_prefix_to_advisor = {}
    for row in student_info:
        advisors = [row['Advisor']]
        if _sanitize(row['Advisor2']):
            advisors.append(row['Advisor2'])
        email_prefix_to_advisor[row['NetID']] = ';'.join(advisors)

    for netid, student in students.items():
        email_prefix = netid
        student['Advisor'] = email_prefix_to_advisor[email_prefix]

    return students


def add_in_assigned(students: StudentsType, assigned: AssignedType,
                    names: NamesType) -> StudentsType:
    for netid, course in assigned.items():
        student = {'Name': names[netid], 'Advisor': '', 'NetID': netid,
                   'Favorite': f"{course[:3]} {course[3:]}", 'Good': '',
                   'OK': '', 'Previous': '', 'Bank': '', 'Join': '',
                   'Time': get_date()}
        students[netid] = student
    return students


def parse_names(student_info: StudentInfoType) -> NamesType:
    names = {}
    for student in student_info:
        names[student['NetID']] = f"{student['First']} {student['Last']}"
    return names


def parse_adjusted(students: StudentsType) -> AdjustedType:
    adjusted = {}
    for netid, student in students.items():
        adjustments = set()
        if adjustments:
            adjusted[netid] = adjustments
    return adjusted


def parse_notes(student_info: StudentInfoType) -> NotesType:
    notes = {}
    for student in student_info:
        notes[student['NetID']] = student['Shorthand']
    return notes


def add_in_notes(students: StudentsType,
                 student_notes: NotesType) -> StudentsType:
    for netid, info in students.items():
        info['Notes'] = student_notes[netid]
        students[netid] = info
    return students


def add_in_bank_join(students: StudentsType,
                     student_info: StudentInfoType) -> StudentsType:
    bank = {}
    join = {}
    for student in student_info:
        netid = student['NetID']
        bank[netid] = student['Bank']
        join[netid] = student['Join']

    for netid, student in students.items():
        student['Bank'] = bank[netid]
        student['Join'] = join[netid]
    return students


def get_students(student_info_sheet_id: str,
                 student_preferences_sheet_id: str) -> tuple[
    AssignedType, YearsType, StudentsType]:
    student_info = get_rows_with_tab_title(
        student_info_sheet_id, gs_consts.PLANNING_INPUT_STUDENTS_TAB_TITLE)
    student_preferences_tab = get_rows_with_tab_title(
        student_preferences_sheet_id, gs_consts.PREFERENCES_INPUT_TAB_TITLE)

    for row in student_info:
        if re.search(r"[mM]issing [Ff]orm", row['Last']):
            student_info.remove(row)
    names = parse_names(student_info)
    student_notes = parse_notes(student_info)
    assigned = parse_pre_assign(student_info)
    years = parse_years(student_info)
    students = parse_student_preferences(student_preferences_tab)
    students = add_in_bank_join(students, student_info)
    # in case assigned students did not send in preferences
    students = add_in_assigned(students, assigned, names)
    students = add_in_notes(students, student_notes)
    students = fix_advisors(students, student_info)
    return assigned, years, students


def get_courses(planning_sheet_id: str) -> CoursesType:
    tab = get_rows_with_tab_title(
        planning_sheet_id, gs_consts.PLANNING_INPUT_COURSES_TAB_TITLE)
    courses = parse_courses(tab)
    return courses


def get_fac_prefs(instructor_preferences_sheet_id: str) -> FacultyPrefsType:
    tab = get_rows_with_tab_title(
        instructor_preferences_sheet_id, gs_consts.PREFERENCES_INPUT_TAB_TITLE)
    prefs = parse_faculty_prefs(tab)
    return prefs


def write_csv(file_name: str, data: list[list]):
    with open(file_name, "w") as file:
        writer = csv.writer(file)
        writer.writerows(data)


def format_pref_list(pref: str) -> str:
    netids = []
    # if they include parens. .*? means shortest match.
    pref = re.sub(r'\(.*?\)', '', pref)
    parts = pref.split(',')
    for part in parts:
        part = re.sub(r'.*<', '', part)  # netid starts after open bracket
        part = re.sub(r'@.*', '', part)  # netid ends at @ sign
        if part:
            netids.append(part)
    netids = ';'.join(netids)
    return netids


def format_course(course: CourseValue, prefs) -> Optional[list[str]]:
    # cols = ['Course','TAs','Weight','Instructor','Title','Notes']
    num = str(course['Course'])
    slots = course['TAs']
    if int(slots) == 0:
        return None
    weight = '' if 'Weight' not in course else course['Weight']
    title = course['Title']
    instructors = ';'.join(re.split(r'[;,]', course['Instructor']))

    course_code = f'COS {num}'
    if re.search(r'[a-zA-Z]', num):
        if 'COS' in num:
            num = num.replace('COS ', '').replace('COS', '')
        else:  # if course num is from another department
            course_code = num.replace(' ', '')

    if course_code in prefs:
        pref = prefs[course_code]
        fav = pref['Favorite']
        veto = pref['Veto']
        fav = format_pref_list(fav)
        veto = format_pref_list(veto)
    else:
        fav = ''
        veto = ''

    course_code = course_code.replace(' ', '')
    return [course_code, slots, weight, fav, veto, instructors, title]


def parse_previous_list(prev: str) -> list[str]:
    if '(' in prev:
        parts = re.split(r"\)[,;]\s?", prev)
    else:
        parts = re.split(r"[,;]\s?", prev)
    regex = re.compile(
        r"^(COS|EGR|PNI)[\s]?([1-9][0-9]{2}[A-F]?)\/?" + r"(COS|EGR|PNI)?[\s]?([1-9][0-9]{2}[A-F]?)?\/?" * 2)

    course_nums = []
    for part in parts:
        match = regex.match(part)
        if match is None:
            continue
        deps_or_nums = match.groups()
        prev_non_cos_dept = False
        for i in range(len(deps_or_nums)):
            matched = deps_or_nums[i]
            if matched is None or "COS" in matched:
                continue
            prev_dept = ''
            if prev_non_cos_dept:
                prev_dept = course_nums.pop()
            prev_non_cos_dept = re.search(r"^[A-Z]", matched)
            if _sanitize(matched):
                course_nums.append(prev_dept + matched)
    return course_nums


def format_prev(prev: str, courses: CoursesType):
    # people didn't follow directions
    prev = prev.replace('),', ');').replace(')\n', ');')
    course_nums = set(parse_previous_list(prev))
    to_delete = set()
    for course in course_nums:
        if re.search(r"[A-Z]", course):
            if course not in courses:
                to_delete.add(course)
        elif int(course) not in courses:
            to_delete.add(course)

    for c in to_delete:
        course_nums.remove(c)

    if not len(course_nums):
        return ''
    course_nums = add_COS(';'.join(course_nums))
    return course_nums


def format_netid(email: str) -> str:
    netid = email.replace('@princeton.edu', '')
    return netid


def format_course_list(courses: str) -> str:
    courses = courses.replace(',', ';')
    courses = courses.replace(' ', '')
    return courses


def format_student(student: StudentType, courses: CoursesType,
                   years: YearsType) -> list[str]:
    # ['Netid','Name','Year','Bank','Join','Weight','Previous','Advisor','Favorite','Good','OK','Notes']
    netid = student['NetID']
    full_name = student['Name']
    year = years[netid]
    bank = student['Bank']
    join = student['Join']
    prev = student['Previous']
    adv = student['Advisor'].replace(',', ';')
    fav = student['Favorite']
    good = student['Good']
    okay = student['OK']
    notes = student['Notes'].replace('"', '\'')
    prev = format_prev(prev, courses)
    row = [netid, full_name, year, bank, join, "", prev, adv, fav, good, okay,
           notes]
    return row


def format_phd(student: StudentType, years: YearsType) -> list[str]:
    # phds = 'Netid,Name,Year,Advisor\n'
    netid = student['NetID']
    full_name = student['Name']
    advisor = student['Advisor'].replace(',', ';')
    year = years.get(netid)
    if not year.strip() or 'PHD' not in year:
        return []
    year = year.replace('PHD', '')
    row = [netid, full_name, year, advisor]
    return row


def format_assigned(netid:str, full_name:str, year:str, advisor:str, course:str, notes:str):
    # ['NetID','Name','Year','Bank','Join','Weight','Previous','Advisor','Favorite','Good','OK','Notes']
    return [netid, full_name, year, "", "", "", "", advisor, course, "", "",
            notes]


def get_date():
    now = datetime.datetime.now()
    date = now.strftime("%Y-%m-%d.%H.%M.%S")
    return date


def make_path(dir_title: str = None):
    if not dir_title:
        dir_title = get_date()
    path = f'data/{dir_title}/inputs'
    os.makedirs(path, exist_ok=True)
    return path


def write_courses(path: str, courses: CoursesType,
                  instructor_prefs: FacultyPrefsType):
    data = [['Course', 'Slots', 'Weight', 'Favorite', 'Veto', 'Instructor',
             'Title']]
    for num in courses:
        course = courses[num]
        row = format_course(course, instructor_prefs)
        if row:
            data.append(row)
    write_csv(f"{path}/course_data.csv", data)


def write_students(path:str, courses:CoursesType, assigned:AssignedType, years:YearsType,
                   students: StudentsType):
    data = [['Netid', 'Name', 'Year', 'Bank', 'Join', 'Weight', 'Previous',
             'Advisors', 'Favorite', 'Good', 'Okay', 'Notes']]
    phds = [['Netid', 'Name', 'Year', 'Advisor']]
    for netid in students:
        if netid in assigned:
            continue
        student = students[netid]
        data.append(format_student(student, courses, years))
        phd_row = format_phd(student, years)
        if phd_row:
            phds.append(phd_row)
    for netid, course in assigned.items():
        student = students[netid]
        data.append(
            format_assigned(
                netid, student['Name'], years[netid], student['Advisor'],
                course, student['Notes']))
    write_csv(f"{path}/student_data.csv", data)
    write_csv(f"{path}/phds.csv", phds)


def write_assigned(path:str, assigned:AssignedType):
    data = [['Netid', 'Course']]
    for netid, course in assigned.items():
        data.append([netid, course])
    write_csv(f"{path}/fixed.csv", data)


def write_adjusted(path: str, adjusted: AdjustedType):
    data = [['Netid', 'Course', 'Weight']]
    for netid, adjustments in adjusted.items():
        for course, weight in adjustments:
            data.append([netid, course, weight])
    write_csv(f"{path}/adjusted.csv", data)


def validate_bank_join_values(students: StudentsType, years: YearsType):
    for netid, student in students.items():
        if student['Bank'] and student['Join']:
            print(f"{netid} has both a bank and a join entry")
        if 'MSE' in years[netid] and (student['Bank'] or student['Join']):
            print(f"{netid} is an MSE student with a bank or a join entry")


def write_csvs(planning_sheet_id: str, student_prefs_sheet_id: str,
               instructor_prefs_sheet_id: str,
               output_directory_title: str = None):
    fac_prefs = get_fac_prefs(instructor_prefs_sheet_id)
    courses = get_courses(planning_sheet_id)
    assigned, years, students = get_students(
        planning_sheet_id, student_prefs_sheet_id)
    validate_bank_join_values(students, years)
    path = make_path(output_directory_title)
    write_courses(path, courses, fac_prefs)
    write_students(path, courses, assigned, years, students)
    write_assigned(path, assigned)
    return planning_sheet_id, student_prefs_sheet_id, instructor_prefs_sheet_id


def check_if_preprocessing_equal(name1: str, name2: str):
    with open(name1) as f1, open(name2) as f2:
        reader1 = list(csv.reader(f1, delimiter=',', quotechar='"'))
        reader2 = list(csv.reader(f2, delimiter=',', quotechar='"'))
        reader1 = sorted(reader1)
        reader2 = sorted(reader2)
        if len(reader1) != len(reader2):
            print(f"big lens of {name1},{name2} not equal")
            return
        for j in range(len(reader1)):
            row1 = reader1[j]
            row2 = reader2[j]
            if len(row1) != len(row2):
                print("lens not equal")
                print(row1, row2)
                continue

            for i in range(len(row1)):
                v1 = row1[i]
                v2 = row2[i]
                if ";" in v1:
                    ar1 = v1.split(";").sort()
                    ar2 = v2.split(";").sort()
                    if ar1 != ar2:
                        print("not equal")
                else:
                    if v1 != v2:
                        print(f"vals {v1},{v2} not equal")
