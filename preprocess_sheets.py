import os
import re
import datetime

import gspread
from oauth2client.client import GoogleCredentials
from google.colab import auth

import params
import g_sheet_consts as gs_consts


def get_rows_with_tab_title(sheet_id, tab_title):
    gc = gspread.authorize(GoogleCredentials.get_application_default())

    sheet = gc.open_by_key(sheet_id).worksheet(tab_title)
    all_rows = sheet.get_all_records()
    return all_rows


def parse_pre_assign(rows):
    fixed = {}
    for row in rows:
        if _sanitize(row["Course"]):
            fixed[row["NetID"]] = add_COS(row["Course"])
    return fixed


def _sanitize(input):
    return input if type(input) != str else input.strip()


def parse_weights(rows):
    weights = {}
    for row in rows:
        netid = row["NetID"]
        bank = float(row["Bank"]) if _sanitize(
            row["Bank"]) else params.DEFAULT_BANK
        join = float(row["Join"]) if _sanitize(
            row["Join"]) else params.DEFAULT_JOIN
        w = params.BANK_MULTIPLIER * (
                bank - params.DEFAULT_BANK) + params.JOIN_MULTIPLIER * (
                    join - params.DEFAULT_JOIN)
        if w:
            weights[netid] = w
    return weights


def parse_years(rows):
    years = {}
    for row in rows:
        netid = row["NetID"]
        track = row["Track"]
        year = row["Year"]
        years[netid] = track + str(year)
    return years


def add_COS(courses):
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


def parse_student_preferences(rows):
    students = {}
    cols = {'Timestamp': 'Time', 'Email': 'Email', 'Name': 'Name',
            'Advisor': 'Advisor', 'What courses at Princeton': 'Previous',
            'Favorite': 'Favorite', 'Good': 'Good', 'OK Match': 'OK'}

    rows = switch_keys_from_rows(rows, cols, True)
    for row in rows:
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
        students[row['Email']] = row
    return students


def parse_courses(rows):
    # cols = ['Course', 'Omit', 'TAs', 'Weight', 'Notes', 'Title']
    courses = {}
    for row in rows:
        courses[row['Course']] = row
    return courses


def switch_keys(dictionary, key_to_switch_dict):
    new = {}
    for key, val in dictionary.items():
        if key in key_to_switch_dict:
            new[key_to_switch_dict[key]] = val
    return new


def switch_keys_from_rows(rows, keys_to_switch_dict, is_paraphrased=False):
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


def parse_fac_prefs(rows):
    cols = {'Timestamp': 'Time', 'Email': 'Email', 'Which course?': 'Course',
            'Best': 'Favorite', 'Avoid': 'Veto'}
    courses = {}
    rows = switch_keys_from_rows(rows, cols, True)
    for row in rows:
        courses[row['Course']] = row
    return courses


def fix_advisors(students, student_info):
    email_prefix_to_advisor = {}
    for row in student_info:
        advisors = [row['Advisor']]
        if _sanitize(row['Advisor2']):
            advisors.append(row['Advisor2'])
        email_prefix_to_advisor[row['NetID']] = ';'.join(advisors)

    for email, student in students.items():
        email_prefix = format_netid(email)
        advisor = email_prefix_to_advisor[email_prefix]
        student['Advisor'] = advisor

    return students


def add_in_assigned(students, assigned, names):
    for netid, course in assigned.items():
        student = {'Name': names[netid], 'Advisor': '',
                   'Email': netid + '@princeton.edu',
                   'Favorite': f"{course[:3]} {course[3:]}", 'Good': '',
                   'OK': '', 'Previous': '', 'Time': get_date()}
        students[netid + '@princeton.edu'] = student
    return students


def parse_names(student_info):
    names = {}
    for student in student_info:
        names[student['NetID']] = f"{student['First']} {student['Last']}"
    return names


def parse_adjusted(students):
    adjusted = {}
    for email, student in students.items():
        netid = format_netid(email)
        adjustments = set()
        for course in re.split(r"[,;]\s?", student['OK']):
            if _sanitize(course):
                adjustments.add((course, params.OKAY_COURSE_PENALTY))
        if adjustments:
            adjusted[netid] = adjustments
    return adjusted


def get_students(student_info_sheet_id, student_preferences_sheet_id):
    student_info = get_rows_with_tab_title(student_info_sheet_id,
                                           gs_consts.PLANNING_INPUT_STUDENTS_TAB_TITLE)
    student_preferences_tab = get_rows_with_tab_title(
        student_preferences_sheet_id, gs_consts.PREFERENCES_INPUT_TAB_TITLE)

    for row in student_info:
        if re.search(r"[mM]issing [Ff]orm", row['Last']):
            student_info.remove(row)
    names = parse_names(student_info)
    assigned = parse_pre_assign(student_info)
    weights = parse_weights(student_info)
    years = parse_years(student_info)
    students = parse_student_preferences(student_preferences_tab)
    # in case assigned students did not send in preferences
    students = add_in_assigned(students, assigned, names)
    students = fix_advisors(students, student_info)
    adjusted = parse_adjusted(students)
    return assigned, weights, years, students, adjusted


def get_courses(planning_sheet_id):
    tab = get_rows_with_tab_title(planning_sheet_id,
                                  gs_consts.PLANNING_INPUT_COURSES_TAB_TITLE)
    courses = parse_courses(tab)
    return courses


def get_fac_prefs(instructor_preferences_sheet_id):
    tab = get_rows_with_tab_title(instructor_preferences_sheet_id,
                                  gs_consts.PREFERENCES_INPUT_TAB_TITLE)
    prefs = parse_fac_prefs(tab)
    return prefs


def write_csv(fname, data):
    f = open(fname, "w")
    f.write(data)
    f.close()


def format_pref_list(pref):
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


def format_course(course, prefs):
    # cols = ['Course','Omit','TAs','Weight','Notes','Title']
    if 'Omit' in course and course['Omit']:
        return ''
    num = str(course['Course'])
    slots = course['TAs']
    if int(slots) == 0:
        return ''
    weight = '' if 'Weight' not in course else course['Weight']
    title = course['Title']

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
    row = f'{course_code},{slots},{weight},{fav},{veto},"{title}"\n'
    return row


def parse_previous_list(prev):
    if '(' in prev:
        parts = re.split(r"\)[,;]\s?", prev)
    else:
        parts = re.split(r"[,;]\s?", prev)
    regex = re.compile(
        r"^(COS|EGR|PNI)[\s]?([1-9][0-9]{2}[A-F]?)\/?" + r"(COS|EGR|PNI)?[\s]?([1-9][0-9]{2}[A-F]?)?\/?" * 2)

    coursenums = []
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
                prev_dept = coursenums.pop()
            prev_non_cos_dept = re.search(r"^[A-Z]", matched)
            if _sanitize(matched):
                coursenums.append(prev_dept + matched)
    return coursenums


def format_prev(prev, courses):
    # """ Previously TA'ing in a different subject that's not COS is not recognized """
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


def format_netid(email):
    netid = email.replace('@princeton.edu', '')
    return netid


def format_course_list(courses):
    courses = courses.replace(',', ';')
    courses = courses.replace(' ', '')
    return courses


def lookup_weight(netid, weights, years):
    weight = 0.0
    if netid in weights:
        weight += weights[netid]
    if netid in years:
        year = years[netid]
        if 'MSE' in year:
            weight += params.MSE_BOOST
    return str(weight) if weight != 0.0 else ''


def format_student(student, courses, weights, years):
    # ['NetID','Name','Weight','Previous','Advisor','Favorite','Good','OK']
    netid = format_netid(student['Email'])
    full_name = student['Name']
    prev = student['Previous']
    adv = student['Advisor'].replace(',', ';')
    fav = student['Favorite']
    good = student['Good']
    okay = student['OK']
    prev = format_prev(prev, courses)
    weight = lookup_weight(netid, weights, years)
    row = f'{netid},{full_name},{weight},{prev},{adv},{fav},{good},{okay}\n'
    return row


def format_phd(student, years):
    # phds = 'Netid,Name,Year,Advisor\n'
    netid = format_netid(student['Email'])
    full_name = student['Name']
    advisor = student['Advisor'].replace(',', ';')
    year = years.get(netid)
    if not year.strip() or 'PHD' not in year:
        return ''
    year = year.replace('PHD', '')
    row = f'{netid},{full_name},{year},{advisor}\n'
    return row


def format_assigned(netid, full_name, advisor, course):
    # ['NetID','Name','Weight','Previous','Advisor','Favorite','Good','OK']
    return f'{netid},{full_name},,,{advisor},{course},,\n'


def get_date():
    now = datetime.datetime.now()
    date = now.strftime("%Y-%m-%d.%H.%M.%S")
    return date


def make_path(dir_title=None):
    if not dir_title:
        dir_title = get_date()
    path = f'data/{dir_title}/inputs'
    os.makedirs(path, exist_ok=True)
    return path


def write_courses(path, courses, prefs):
    data = 'Course,Slots,Weight,Favorite,Veto,Title\n'
    for num in courses:
        course = courses[num]
        data += format_course(course, prefs)
    write_csv(f"{path}/course_data.csv", data)


def write_students(path, courses, assigned, weights, years, students):
    data = 'Netid,Name,Weight,Previous,Advisors,Favorite,Good,Okay\n'
    phds = 'Netid,Name,Year,Advisor\n'
    for email in students:
        if format_netid(email) in assigned:
            continue
        student = students[email]
        data += format_student(student, courses, weights, years)
        phds += format_phd(student, years)
    for netid, course in assigned.items():
        student = students[netid + '@princeton.edu']
        data += format_assigned(netid, student['Name'], student['Advisor'],
                                course)
    write_csv(f"{path}/student_data.csv", data)
    write_csv(f"{path}/phds.csv", phds)


def write_assigned(path, assigned):
    data = 'Netid,Course\n'
    for netid, course in assigned.items():
        data += f"{netid},{course}\n"
    write_csv(f"{path}/fixed.csv", data)


def write_adjusted(path, adjusted):
    data = 'Netid,Course,Weight\n'
    for netid, adjustments in adjusted.items():
        for course, weight in adjustments:
            data += f"{netid},{course},{weight}\n"
    write_csv(f"{path}/adjusted.csv", data)


def write_csvs(output_directory_title=None, planning_sheet_id=None,
               student_prefs_sheet_id=None, instructor_prefs_sheet_id=None):
    auth.authenticate_user()

    if not planning_sheet_id:
        planning_sheet_id = '1xcq575Wp_JMQ0j-6udK_Lwq3BrSZnL-ro_UZagIhAGk'
    if not student_prefs_sheet_id:
        student_prefs_sheet_id = '1TX6kEsIYB2rPqFnYkXyZbzN8CaK-jEh29Si40BWgCQ8'
    if not instructor_prefs_sheet_id:
        instructor_prefs_sheet_id = '1poN30Xdgs4_4STw-GxQuqE3UAq3WuNq02EqknJZScFo'

    fac_prefs = get_fac_prefs(instructor_prefs_sheet_id)
    courses = get_courses(planning_sheet_id)
    assigned, weights, years, students, adjusted = get_students(
        planning_sheet_id, student_prefs_sheet_id)
    path = make_path(output_directory_title)
    write_courses(path, courses, fac_prefs)
    write_students(path, courses, assigned, weights, years, students)
    write_assigned(path, assigned)
    write_adjusted(path, adjusted)
    return planning_sheet_id, student_prefs_sheet_id, instructor_prefs_sheet_id


if __name__ == '__main__':
    write_csvs()
