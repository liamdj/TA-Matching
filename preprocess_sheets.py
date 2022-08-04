import argparse
import csv
import datetime
import os
import re
from typing import Union, Any, Optional, List, Dict, Tuple, Set

import g_sheet_consts as gs_consts
import write_to_google_sheets as write_gs

try:
    from typing import TypedDict
except ImportError:
    from typing_extensions import TypedDict


class StudentType(TypedDict, total=False):
    NetID: str
    Name: str
    Advisor: str
    Previous: str
    Favorite: str
    Good: str
    OK: str
    Sorted: str
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
    Sorted: str


StudentsType = Dict[str, StudentType]
StudentInfoType = List[StudentInfoValue]
CoursesType = Dict[str, CourseValue]
FacultyPrefsType = Dict[str, FacultyPref]
AdjustedType = Dict[str, Set[str]]
AssignedType = Dict[str, str]
NamesType = Dict[str, str]
NotesType = Dict[str, str]
YearsType = Dict[str, str]


def get_rows_with_tab_title(sheet_id: str, tab_title: str) -> Union[
    list, List[dict]]:
    sheet = write_gs.get_sheet_by_id(sheet_id).worksheet(tab_title)
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
        student_prefs: List[Dict[str, str]]) -> StudentsType:
    students = {}
    cols = {'Email': 'NetID', 'Name': 'Name', 'Advisor': 'Advisor',
            'previously': 'Previous', 'Favorite': 'Favorite',
            'Good': 'Good', 'OK Match': 'OK', 'Sorted': 'Sorted'}

    student_prefs = switch_keys_from_rows(student_prefs, cols, True)
    for row in student_prefs:
        fav = format_course_list(row['Favorite'])
        fav_copy = []
        for fav_c in fav:
            if fav_c not in fav_copy:
                fav_copy.append(fav_c)
        fav = fav_copy
        good = set(format_course_list(row['Good']))
        okay = set(format_course_list(row['OK']))

        for fav_c in fav:
            if fav_c in good:
                good.remove(fav_c)
            if fav_c in okay:
                okay.remove(fav_c)
        for good_c in good:
            if good_c in okay:
                okay.remove(good_c)

        row['Name'] = row['Name'].strip()
        row['Sorted'] = parse_sorted_favorite_list(row.get('Sorted', ''), fav)
        row['Favorite'] = ';'.join(fav)
        row['Good'] = ';'.join(good)
        row['OK'] = ';'.join(okay)
        row['NetID'] = format_netid(row['NetID'])
        students[row['NetID']] = row
    return students


def parse_sorted_favorite_list(sorted_favorites: str,
                               favorite_courses: List[str]) -> str:
    sorted_favorites = sorted_favorites.replace(',', '>')
    sorted_favorites = sorted_favorites.replace(';', '>')
    sorted_favorites = sorted_favorites.replace(' ', '')
    sorted_favorites = sorted_favorites.replace('\n', '')
    sorted_favorites = sorted_favorites.split('>')
    all_courses_in_sorted_favs = []
    for courses in sorted_favorites:
        courses_set = set()
        for fav_c in courses.split('='):
            if fav_c in favorite_courses:
                courses_set.add(fav_c)
        if len(courses_set):
            all_courses_in_sorted_favs.append('='.join(courses_set))
    return '>'.join(all_courses_in_sorted_favs)


def parse_courses(course_info: List[dict]) -> CoursesType:
    courses = {}
    for row in course_info:
        courses[row['Course']] = row
    return courses


def switch_keys(dictionary: Dict[str, Any],
                key_to_switch_dict: Dict[str, str]) -> Dict[str, Any]:
    new = {}
    for key, val in dictionary.items():
        if key in key_to_switch_dict:
            new[key_to_switch_dict[key]] = val
    return new


def switch_keys_from_rows(rows: List[dict], keys_to_switch_dict: Dict[str, str],
                          is_paraphrased=False) -> List[dict]:
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


def fix_advisors(students: StudentsType,
                 student_info: StudentInfoType) -> StudentsType:
    email_prefix_to_advisor = {}
    for row in student_info:
        advisors = [row['Advisor']]
        if _sanitize(row['Advisor2']):
            advisors.append(row['Advisor2'])
        email_prefix_to_advisor[row['NetID']] = ';'.join(advisors)

    for netid, student in students.items():
        if netid not in email_prefix_to_advisor:
            continue
        email_prefix = netid
        student['Advisor'] = email_prefix_to_advisor[email_prefix]

    return students


def add_in_assigned(students: StudentsType, assigned: AssignedType,
                    names: NamesType) -> StudentsType:
    for netid, course in assigned.items():
        student = {'Name': names[netid], 'Advisor': '', 'NetID': netid,
                   'Favorite': f"{course[:3]} {course[3:]}", 'Good': '',
                   'OK': '', 'Previous': '', 'Bank': '', 'Join': '',
                   'Time': get_date(), 'Sorted': ''}
        students[netid] = student
    return students


def parse_names(student_info: StudentInfoType) -> NamesType:
    names = {}
    for student in student_info:
        names[student[
            'NetID']] = f"{student['First'].strip()} {student['Last'].strip()}"
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
        if 'Shorthand' in student:
            notes[student['NetID']] = student['Shorthand']
    return notes


def add_in_notes(students: StudentsType,
                 student_notes: NotesType) -> StudentsType:
    for netid, info in students.items():
        if netid not in student_notes:
            continue
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
        if netid not in bank or netid not in join:
            print('Student with netid', netid, 'found in student form responses but not in planning spreadsheet')
            continue
        student['Bank'] = bank[netid]
        student['Join'] = join[netid]
    return students


def get_students(planning_sheet_worksheets: List[write_gs.Worksheet],
                 student_preferences_sheet_id: str) -> Tuple[
    AssignedType, YearsType, StudentsType]:
    student_info = get_student_info(planning_sheet_worksheets)
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


def get_student_info(
        planning_sheet_worksheets: List[write_gs.Worksheet]) -> StudentInfoType:
    student_info = write_gs.get_worksheet_from_worksheets(
        planning_sheet_worksheets, gs_consts.PLANNING_INPUT_STUDENTS_TAB_TITLE,
        gs_consts.MATCHING_OUTPUT_SHEET_TITLE).get_all_records()

    for student in student_info:
        student['NetID'] = student['NetID'].strip()
    return student_info


def get_courses(planning_sheet_id: str) -> Tuple[
    CoursesType, str, List[write_gs.Worksheet]]:
    planning_sheet = write_gs.get_sheet_by_id(planning_sheet_id)
    planning_worksheets = planning_sheet.worksheets()
    ws = write_gs.get_worksheet_from_worksheets(
        planning_worksheets, gs_consts.PLANNING_INPUT_COURSES_TAB_TITLE,
        gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    courses = parse_courses(ws.get_all_records())
    return courses, planning_sheet.title, planning_worksheets


def get_fac_prefs(instructor_preferences_sheet_id: str) -> FacultyPrefsType:
    tab = get_rows_with_tab_title(
        instructor_preferences_sheet_id, gs_consts.PREFERENCES_INPUT_TAB_TITLE)
    cols = {'Email': 'Email', 'Which course?': 'Course', 'Best': 'Favorite',
            'Avoid': 'Veto', 'Sorted': 'Sorted'}
    fac_prefs = {}
    switched_keys = switch_keys_from_rows(tab, cols, True)
    for row in switched_keys:
        fac_prefs[row['Course']] = row
    return fac_prefs


def write_csv(file_name: str, data: List[list]):
    with open(file_name, "w") as file:
        writer = csv.writer(file)
        writer.writerows(data)


def format_pref_list(pref: str) -> List[str]:
    netids = []
    # if they include parens. .*? means shortest match.
    pref = re.sub(r'\(.*?\)', '', pref)
    parts = pref.split(',')
    for part in parts:
        part = re.sub(r'.*<', '', part)  # netid starts after open bracket
        part = re.sub(r'@.*', '', part)  # netid ends at @ sign
        if part:
            netids.append(part)
    return netids


def format_course(course: CourseValue, prefs: FacultyPrefsType) -> Optional[
    List[str]]:
    num = str(course['Course'])
    slots = course['TAs']
    if int(slots) == 0:
        return None
    weight = '' if 'Weight' not in course else course['Weight']
    title = course['Title']
    instructors = ';'.join(re.split(r'[;,]', course['Instructor']))

    if re.search(r'[a-zA-Z]', num):
        if 'COS' in num:
            num = num.replace('COS ', '').replace('COS', '')
            course_code = f'COS {num}'
        else:  # if course num is from another department
            course_code = num.replace(' ', '')

    sort_fav = ''
    if course_code in prefs:
        pref = prefs[course_code]
        fav = ';'.join(format_pref_list(pref['Favorite']))
        veto = ';'.join(format_pref_list(pref['Veto']))
        if pref.get('Sorted') and pref['Sorted'].title() != 'False':
            sort_fav = 'Yes'
    else:
        fav = ''
        veto = ''

    course_code = course_code.replace(' ', '')
    return [course_code, slots, weight, fav, sort_fav, veto, instructors, title]


def parse_previous_list(prev: str) -> List[str]:
    if '(' in prev:
        parts = re.split(r"\)[,;]\s?", prev)
    else:
        parts = re.split(r"[,;]\s?", prev)
    regex = re.compile(
        r"^(COS|EGR|PNI|ELE|MAE|ISC)[\s]?([1-9][0-9]{2}[A-F]?)\/?" + r"(COS|EGR|PNI|ELE|MAE|ISC)?[\s]?([1-9][0-9]{2}[A-F]?)?\/?" * 2)

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
    return netid.strip()


def format_course_list(courses: str) -> List[str]:
    courses = courses.replace(',', ';')
    courses = courses.replace(' ', '')
    courses = courses.split(';')
    # for i in range(len(courses)):
    #     courses[i] = re.search('[A-z]{3}?\s[0-9]{3}', courses[i]).group()
    return courses


def format_student(student: StudentType, courses: CoursesType,
                   years: YearsType) -> List[str]:
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
    sorted_favs = student['Sorted']
    notes = student['Notes'].replace('"', '\'') if 'Notes' in student else ''
    prev = format_prev(prev, courses)
    row = [netid, full_name, year, bank, join, "", prev, adv, fav, good, okay,
           sorted_favs, notes]
    return row


def format_phd(student: StudentType, years: YearsType) -> List[str]:
    netid = student['NetID']
    full_name = student['Name']
    advisor = student['Advisor'].replace(',', ';')
    year = years.get(netid)
    if not year.strip() or 'PHD' not in year:
        return []
    year = year.replace('PHD', '')
    row = [netid, full_name, year, advisor]
    return row


def format_assigned(netid: str, full_name: str, year: str, advisor: str,
                    course: str, notes: str) -> List[str]:
    return [netid, full_name, year, "", "", "", "", advisor, course, "", "", "",
            notes]


def get_date() -> str:
    now = datetime.datetime.now()
    date = now.strftime("%Y-%m-%d.%H.%M.%S")
    return date


def make_path(dir_title: str = None) -> str:
    if not dir_title:
        dir_title = get_date()
    path = f'data/{dir_title}/inputs'
    os.makedirs(path, exist_ok=True)
    return path


def write_courses(path: str, courses: CoursesType,
                  instructor_prefs: FacultyPrefsType):
    """
    `FacultyPrefsType` includes a field `Sorted` which will be written under
    the column `Favorites Sorted`. If the value in the `Sorted` field is present
    (unless it is `str(False)` or `str(false)`) then it will be interpreted as
    `True`.
    """
    data = [
        ['Course', 'Slots', 'Weight', 'Favorite', 'Favorites Sorted', 'Veto',
         'Instructor', 'Title']]
    for num in courses:
        course = courses[num]
        row = format_course(course, instructor_prefs)
        if row:
            data.append(row)
    write_csv(f"{path}/course_data.csv", data)


def write_students(path: str, courses: CoursesType, assigned: AssignedType,
                   years: YearsType, students: StudentsType):
    data = [['NetID', 'Name', 'Year', 'Bank', 'Join', 'Weight', 'Previous',
             'Advisors', 'Favorite', 'Good', 'Okay', 'Sorted Favorites',
             'Notes']]
    phds = [['NetID', 'Name', 'Year', 'Advisor']]
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
                course, student['Notes'] if 'Notes' in student else ''))
    write_csv(f"{path}/student_data.csv", data)
    write_csv(f"{path}/phds.csv", phds)


def write_assigned(path: str, assigned: AssignedType):
    data = [['NetID', 'Course']]
    for netid, course in assigned.items():
        data.append([netid, course])
    write_csv(f"{path}/fixed.csv", data)


def write_adjusted(path: str, adjusted: AdjustedType):
    data = [['NetID', 'Course', 'Weight']]
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


def get_and_write_previous(path: str, previous_matching_ws_title: str):
    matching_sheet = write_gs.get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    previous_ws = write_gs.get_worksheet_from_sheet(
        matching_sheet, previous_matching_ws_title)
    values = previous_ws.get_all_values()
    with open(path + '/previous.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerows(values)


def write_csvs(planning_sheet_id: str, student_prefs_sheet_id: str,
               instructor_prefs_sheet_id: str,
               previous_matching_ws_title: str = None,
               output_directory_title: str = None) -> Tuple[
    str, str, str, str, List[write_gs.Worksheet]]:
    fac_prefs = get_fac_prefs(instructor_prefs_sheet_id)
    courses, planning_sheet_title, planning_sheet_worksheets = get_courses(
        planning_sheet_id)
    assigned, years, students = get_students(
        planning_sheet_worksheets, student_prefs_sheet_id)
    validate_bank_join_values(students, years)
    path = make_path(output_directory_title)
    write_courses(path, courses, fac_prefs)
    write_students(path, courses, assigned, years, students)
    write_assigned(path, assigned)
    if previous_matching_ws_title:
        get_and_write_previous(path, previous_matching_ws_title)
    return planning_sheet_id, student_prefs_sheet_id, instructor_prefs_sheet_id, planning_sheet_title, planning_sheet_worksheets


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'planning', type=str,
        help='The Google Sheets id for the planning sheet')
    parser.add_argument(
        'ta_prefs', type=str,
        help='The Google Sheets id for the student preferences sheet')
    parser.add_argument(
        'fac_prefs', type=str,
        help='The Google Sheets id for the instructor preferences sheet')
    parser.add_argument(
        '--output_dir', type=str, required=False, default='matching',
        help='The name of the directory to which the preprocessed csvs will be saved')
    args = parser.parse_args()
    write_csvs(
        args.planning, args.ta_prefs, args.fac_prefs,
        output_directory_title=args.output_dir)
