import os
import re
import csv
import datetime

import gspread
from oauth2client.client import GoogleCredentials
from google.colab import auth


def get_rows(sheet_id, tabs):
    gc = gspread.authorize(GoogleCredentials.get_application_default())

    sheets = []
    for i in tabs:
        sheet = gc.open_by_key(sheet_id).get_worksheet(i)
        all_rows = sheet.get_all_records()
        # all_rows = sheet.get_all_values()
        # all_rows = all_rows[1:]  # get rid of header row
        sheets.append(all_rows)
    return sheets


def get_rows_with_tab_title(sheet_id, tab_title):
    gc = gspread.authorize(GoogleCredentials.get_application_default())

    sheet = gc.open_by_key(sheet_id).worksheet(tab_title)
    all_rows = sheet.get_all_records()
    # all_rows = sheet.get_all_values()
    # all_rows = all_rows[1:]  # get rid of header row
    return all_rows


def parse_pre_assign(rows):
    fixed = {}
    for row in rows:
        if row["Course"]:
            fixed[row["NetId"]] = add_COS(row["Course"])
        # if row[12]:
        #     fixed[row[3]] = add_COS(row[12])
    return fixed


def parse_weights(rows):
    weights = {}
    for row in rows:
        netid = row["NetId"]
        # netid = row[3]
        bank = float(row["Bank"]) if row["Bank"] else 0.0
        # bank = float(row[9]) if row[9] else 0.0
        join = float(row["Join"]) if row["Join"] else 0.0
        # join = float(row[10]) if row[10] else 0.0
        w = -2*(bank + 1) + 2*(join - 3)
        if w:
            weights[netid] = w
    return weights


def parse_years(rows):
    years = {}
    for row in rows:
        netid = row["NetId"]
        track = row["Track"]
        year = row["Year"]
        # netid = row[3]
        # track = row[5]
        # year = row[6]
        years[netid] = track + str(year)
    return years


def add_COS(courses):
    """ Ignores COS """
    parts = str(courses).split(';')
    for i, part in enumerate(parts):
        # if course code from a non-COS discipline
        if re.search(r'[a-zA-Z]', part):
            parts[i] = part.replace(' ', '')
        else:
            parts[i] = 'COS' + part
    courses = ';'.join(parts)
    return courses


def map_row_to_obj(row, cols):
    obj = {}
    n = len(cols)
    for i in range(n):
        col = cols[i]
        val = row[i]
        obj[col] = val.strip()
    return obj


def student_last_name(student):
    full = student['Name']
    if not full:
        return ''
    parts = full.split()
    if not len(parts):
        return ''
    last = parts[-1]
    return last


def parse_student_preferences(rows):
    students = {}
    cols = {'Timestamp': 'Time', 'Email': 'Email', 'Name': 'Name', 'Advisor': 'Advisor', 'What courses have you previously':
            'Previous', 'Favorite': 'Favorite', 'Good': 'Good', 'OK': 'OK'}

    rows = switch_keys_from_rows(rows, cols, True)
    for row in rows:
        students[row['Email']] = row
    #objects.sort(key=lambda obj: student_last_name(obj))
    return students


def map_rows_to_course_dict(rows, cols):
    courses = {}
    for row in rows:
        obj = map_row_to_obj(row, cols)
        if obj:
            course = obj['Course']
            courses[course] = obj
    return courses


def parse_courses(rows):
    # cols = ['Course', 'Omit', 'TAs', 'Weight', 'Notes', 'Title']
    courses = {}
    for row in rows:
        row['TAs'] = row['Grad TAs needed']
        del row['Grad TAs needed']
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
    If `is_paraphrased`, then assume that the the keys in `keys_to_switch_dict`
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
    cols = {'Timestamp': 'Time', 'Email Address': 'Email', 'Which course?': 'Course',
            'Best Match(es)': 'Favorite', 'Matches to Avoid': 'Veto'}
    courses = {}
    rows = switch_keys_from_rows(rows, cols)
    for row in rows:
        courses[row['Course']] = row
    return courses


def get_students(student_info_sheet_id=None, student_preferences_sheet_id=None):
    if not student_info_sheet_id:
        student_info_sheet_id = '1flSWN5vpzp4hK76mMdn-2D-0niZoDuOpSaNHO4vpeyY'
    if not student_preferences_sheet_id:
        student_preferences_sheet_id = '1dbmB8WMHMtnOoPjwZAH54uxtoj7k9J4qNVWPjAf-XyI'

    student_info = get_rows_with_tab_title(student_info_sheet_id, 'Students')
    student_preferences_tab = get_rows_with_tab_title(
        student_preferences_sheet_id, 'Form Responses 1')

    assigned = parse_pre_assign(student_info)
    weights = parse_weights(student_info)
    years = parse_years(student_info)
    students = parse_student_preferences(student_preferences_tab)
    return assigned, weights, years, students


def get_courses(course_info_sheet_id=None):
    if not course_info_sheet_id:
        course_info_sheet_id = '1Ok7yctDd20l0v0fJ8-iQkwokiB_r1SSLlCPWBxeWYlU'
    tab = get_rows_with_tab_title(course_info_sheet_id, "TA Match Targets")
    courses = parse_courses(tab)
    return courses


def get_fac_prefs(instructor_preferences_sheet_id=None):
    if not instructor_preferences_sheet_id:
        instructor_preferences_sheet_id = '1G0W_Kf3nC4HJWH91joRxmbkfDodAzsFO62elSOahW4U'
    tab = get_rows_with_tab_title(
        instructor_preferences_sheet_id, "Form Responses 1")
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
    if course['Omit']:
        return ''
    num = str(course['Course'])
    slots = course['TAs']
    weight = course['Weight']
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


def format_prev(prev, courses):
    """ Previously TA'ing in a different subject that's not COS is not recognized """
    prev = prev.replace('),', ');').replace(
        ')\n', ');')  # people didn't follow directions
    parts = prev.split(';')
    coursenums = []
    for part in parts:
        parens = part.split('(')
        if len(parens) < 2 or 'Princeton' not in parens[1]:
            continue
        numbers = re.split(r'\D+', parens[0])
        for num in numbers:
            if num and int(num) in courses:
                coursenums.append(num)
    if not len(coursenums):
        return ''
    coursenums = set(coursenums)  # presumably to remove duplicates
    coursenums = list(coursenums)
    coursenums = ';'.join(coursenums)
    coursenums = add_COS(coursenums)
    return coursenums


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
            weight += 20.0
    return str(weight) if weight != 0.0 else ''


def lookup_year(netid, years):
    if netid not in years:
        return ''
    year = years[netid]
    if 'G2' in year:
        return 'G2'
    return ''


def format_student(student, courses, weights, years):
    # ['NetID','Name','Weight','Previous','Advisor','Favorite','Good','OK']
    netid = format_netid(student['Email'])
    full_name = student['Name']
    prev = student['Previous']
    adv = student['Advisor'].replace(',', ';')
    fav = student['Favorite']
    good = student['Good']
    okay = student['OK']
    weight = lookup_weight(netid, weights, years)
    prev = format_prev(prev, courses)
    fav = format_course_list(fav)
    good = format_course_list(good)
    okay = format_course_list(okay)
    row = f'{netid},{full_name},{weight},{prev},{adv},{fav},{good},{okay}\n'
    return row


def format_phd(student, years):
    # phds = 'Netid,Name,Year,Advisor\n'
    netid = format_netid(student['Email'])
    full_name = student['Name']
    advisor = student['Advisor'].replace(',', ';')
    year = years.get(netid)
    if not year or 'PHD' not in year:
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


def make_path():
    date = get_date()
    path = f'data/{date}/inputs'
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
        data += format_assigned(netid,
                                student['Name'], student['Advisor'], course)
    write_csv(f"{path}/student_data.csv", data)
    write_csv(f"{path}/phds.csv", phds)


def write_assigned(path, assigned):
    data = 'Netid,Course\n'
    for netid, course in assigned.items():
        data += f"{netid},{course}\n"
    write_csv(f"{path}/fixed.csv", data)


'''
def get_advisors(student_info_sheet_id=None):
    if not student_info_sheet_id:
        student_info_sheet_id = '1flSWN5vpzp4hK76mMdn-2D-0niZoDuOpSaNHO4vpeyY'
    rows = get_rows_with_tab_title(student_info_sheet_id, 'Advisors')
    return parse_advisors(rows)


def parse_advisors(rows):
    adv = {}
    # Advisor, Course
    for row in rows:
        if len(row) < 2:
            continue
        a = row[0].strip()
        c = row[1].strip()
        if a and c:
            adv[a] = add_COS(c)
    return adv



def student_has_course(student, course):
    courses = student['Courses']
    return (course in courses)

def student_course_pref(student, course):
    for pref in ['Favorite','Good','OK']:
        if course in student[pref]:
            return pref
    return 'Bad'
'''


def write_csvs(student_info_sheet_id=None, student_preferences_sheet_id=None, instructor_preferences_sheet_id=None, course_info_sheet_id=None):
    auth.authenticate_user()
    fac_prefs = get_fac_prefs(instructor_preferences_sheet_id)
    courses = get_courses(course_info_sheet_id)
    assigned, weights, years, students = get_students(
        student_info_sheet_id, student_preferences_sheet_id)
    path = make_path()
    write_courses(path, courses, fac_prefs)
    write_students(path, courses, assigned, weights, years, students)
    write_assigned(path, assigned)


if __name__ == '__main__':
    write_csvs()
    # courses = {109: {'Course': 109, 'Omit': 'x', 'Weight': '', 'Notes': '', 'Title': 'Computers in Our World', 'TAs': 0}, 126: {'Course': 126, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Computer Science: An Interdisciplinary Approach', 'TAs': 10}, 217: {'Course': 217, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Introduction to Programming Systems', 'TAs': 4}, 226: {'Course': 226, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Algorithms and Data Structures', 'TAs': 6}, 240: {'Course': 240, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Reasoning about Computation', 'TAs': 5}, 302: {'Course': 302, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Mathematics for Numerical Computing and Machine Learning', 'TAs': 2}, 316: {'Course': 316, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Principles of Computer System Design', 'TAs': 2}, 324: {'Course': 324, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Introduction to Machine Learning', 'TAs': 9}, 333: {'Course': 333, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Advanced Programming Techniques', 'TAs': 3}, 418: {'Course': 418, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Distributed Systems', 'TAs': 4}, 426: {'Course': 426, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Computer Graphics', 'TAs': 4}, 432: {
    #     'Course': 432, 'Omit': '', 'Weight': 1, 'Notes': '', 'Title': 'Information Security', 'TAs': 5}, 445: {'Course': 445, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Economics and Computing', 'TAs': 4}, 475: {'Course': 475, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Computer Architecture', 'TAs': 1}, 484: {'Course': 484, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Natural Language Processing', 'TAs': 2}, 488: {'Course': 488, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Introduction to Analytic Combinatorics', 'TAs': 3}, 495: {'Course': 495, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Special Topics in Computer Science: Web3: Blockchains, Cryptocurrencies, and Decentralization', 'TAs': 2}, 511: {'Course': 511, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Theoretical Machine Learning', 'TAs': 2}, 518: {'Course': 518, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Advanced Computer Systems', 'TAs': 1}, 522: {'Course': 522, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Computational Complexity', 'TAs': 2}, 585: {'Course': 585, 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Information Theory and Applications', 'TAs': 1}, 'EGR154': {'Course': 'EGR154', 'Omit': '', 'Weight': '', 'Notes': '', 'Title': 'Foundations of Engineering: Linear Systems', 'TAs': 1}}
    # prev = """COS 522 (Spring20, Princeton); COS 511 (Spring21, Princeton); COS 320 (Fall21, Princeton); COS 518 (Spring19, Qduqmqgnb)"""
    # format_prev(prev, courses)
