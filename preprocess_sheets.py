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
        all_rows = sheet.get_all_values()
        all_rows = all_rows[1:]  # get rid of header row
        sheets.append(all_rows)
    return sheets


def parse_pre_assign(rows):
    pre = {}
    for row in rows:
        if len(row) < 4:
            continue
        email = row[1].strip()
        course = row[3].strip()
        if email and course:
            # Name, Email, Advisor, Course
            pre[email] = row[:4]
    return pre


def parse_weights(rows):
    weights = {}
    for row in rows:
        if len(row) < 4:
            continue
        email = row[1].strip()
        if email:
            # Name, Email, Advisor, Weight
            weights[email] = row[3]
    return weights


def parse_years(rows):
    years = {}
    for row in rows:
        if len(row) < 3:
            continue
        email = row[1].strip()
        year = row[2].strip()
        if email and year:
            years[email] = year
    return years


def add_COS(courses):
    parts = courses.split(';')
    for i, part in enumerate(parts):
        parts[i] = 'COS' + part
    courses = ';'.join(parts)
    return courses


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


def parse_students(rows):
    students = {}
    cols = ['Time', 'Email', 'Name', 'Advisor',
            'Previous', 'Favorite', 'Good', 'OK']
    for row in rows:
        obj = map_row_to_obj(row, cols)
        if obj:
            email = obj['Email']
            students[email] = obj
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
    cols = ['Course', 'Omit', 'TAs', 'Weight', 'Notes', 'Title']
    courses = map_rows_to_course_dict(rows, cols)
    return courses


def parse_fac_prefs(rows):
    cols = ['Time', 'Email', 'Course', 'Favorite', 'Veto']
    courses = map_rows_to_course_dict(rows, cols)
    return courses


def get_students():
    G_SHEET_ID = '19TTFTx2mLcsZnat55A_2s0qoYTqXgkusmx2yXmmB1oQ'
    # tabs are: 0:form 1:year 2:weights 3:pre-assign 4:advisor
    tabs = list(range(5))
    tabs = get_rows(G_SHEET_ID, tabs)
    advisors = parse_advisors(tabs[4])
    assigned = parse_pre_assign(tabs[3])
    weights = parse_weights(tabs[2])
    years = parse_years(tabs[1])
    students = parse_students(tabs[0])
    return advisors, assigned, weights, years, students


def get_courses():
    G_SHEET_ID = '1pJwncs-qoLMFjYoiY4AhVp1TnhrPsDIWMVtUsApsAG0'
    tabs = [1]  # targets
    tabs = get_rows(G_SHEET_ID, tabs)
    courses = parse_courses(tabs[0])
    return courses


def get_fac_prefs():
    G_SHEET_ID = '1kjL5u8LD9PDTXo2fyIMpMnXIiZbL9DbZoVUgeGUQotU'
    tabs = [0]  # form answers
    tabs = get_rows(G_SHEET_ID, tabs)
    prefs = parse_fac_prefs(tabs[0])
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
    omit = course['Omit']
    num = course['Course']
    slots = course['TAs']
    weight = course['Weight']
    title = course['Title']
    cosnum = f'COS {num}'
    if cosnum in prefs:
        pref = prefs[cosnum]
        fav = pref['Favorite']
        veto = pref['Veto']
        fav = format_pref_list(fav)
        veto = format_pref_list(veto)
    else:
        fav = ''
        veto = ''
    row = f'COS{num},{slots},{weight},{fav},{veto},{title}\n'
    if omit == 'x':
        return ''
    return row


def format_prev(prev, courses):
    prev = prev.replace('),', ');')  # people didn't follow directions
    parts = prev.split(';')
    coursenums = []
    for part in parts:
        parens = part.split('(')
        if len(parens) < 2 or 'Princeton' not in parens[1]:
            continue
        numbers = re.split(r'\D+', parens[0])
        for num in numbers:
            if num in courses:
                coursenums.append(num)
    if not len(coursenums):
        return ''
    coursenums = set(coursenums)
    coursenums = list(coursenums)
    coursenums = ';'.join(coursenums)
    coursenums = add_COS(coursenums)
    return coursenums


def format_netid(email):
    netid = email.replace('@princeton.edu', '')
    return netid


def format_advisor(adv, advisors):
    if adv in advisors:
        adv = advisors[adv]
    else:
        adv = ''
    return adv


def format_course_list(courses):
    courses = courses.replace(',', ';')
    courses = courses.replace(' ', '')
    return courses


def lookup_weight(netid, weights, years):
    if netid in weights:
        return weights[netid]
    elif netid in years:
        year = years[netid]
        if 'MSE' in year:
            return '20'
    return ''


def lookup_year(netid, years):
    if netid not in years:
        return ''
    year = years[netid]
    if 'G2' in year:
        return 'G2'
    return ''


def format_student(student, advisors, courses, weights, years):
    # ['Time','Email','Name','Advisor','Previous','Favorite','Good','OK']
    netid = student['Email']
    full = student['Name']
    prev = student['Previous']
    adv = student['Advisor']
    fav = student['Favorite']
    good = student['Good']
    okay = student['OK']
    weight = lookup_weight(netid, weights, years)
    prev = format_prev(prev, courses)
    netid = format_netid(netid)
    adv = format_advisor(adv, advisors)
    fav = format_course_list(fav)
    good = format_course_list(good)
    okay = format_course_list(okay)
    row = f'{netid},{full},{weight},{prev},{adv},{fav},{good},{okay}\n'
    return row


def format_phd(student, years):
    # phds = 'Netid,Name,Year,Advisor\n'
    netid = student['Email']
    full = student['Name']
    adv = student['Advisor']
    year = lookup_year(netid, years)
    if not year:
        return ''
    row = f'{netid},{full},{year},{adv}\n'
    return row


def format_assigned(student):
    # ['Time','Email','Name','Advisor','Previous','Favorite','Good','OK']
    netid = student[1]
    full = student[0]
    course = student[3]
    netid = format_netid(netid)
    course = add_COS(course)
    row = f'{netid},{full},,,,{course},,\n'
    return row


def format_assignment(student):
    # Netid,Course
    netid = student[1]
    course = student[3]
    netid = format_netid(netid)
    course = add_COS(course)
    row = f'{netid},{course}\n'
    return row


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


def write_students(path, courses, advisors, assigned, weights, years, students):
    data = 'Netid,Name,Weight,Previous,Advisors,Favorite,Good,Okay\n'
    phds = 'Netid,Name,Year,Advisor\n'
    for email in students:
        if email in assigned:
            continue
        student = students[email]
        data += format_student(student, advisors, courses, weights, years)
        phds += format_phd(student, years)
    for email in assigned:
        student = assigned[email]
        data += format_assigned(student)
    write_csv(f"{path}/student_data.csv", data)
    write_csv(f"{path}/phds.csv", phds)


def write_assigned(path, assigned):
    data = 'Netid,Course\n'
    for email in assigned:
        student = assigned[email]
        data += format_assignment(student)
    write_csv(f"{path}/fixed.csv", data)


'''
def student_has_course(student, course):
    courses = student['Courses']
    return (course in courses)

def student_course_pref(student, course):
    for pref in ['Favorite','Good','OK']:
        if course in student[pref]:
            return pref
    return 'Bad'
'''


def main():
    auth.authenticate_user()
    prefs = get_fac_prefs()
    courses = get_courses()
    advisors, assigned, weights, years, students = get_students()
    path = make_path()
    write_courses(path, courses, prefs)
    write_students(path, courses, advisors, assigned, weights, years, students)
    write_assigned(path, assigned)


if __name__ == '__main__':
    main()
