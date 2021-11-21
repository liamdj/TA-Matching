import datetime
import random
import string

COURSES = ['COS 126', 'COS 217', 'COS 302', 'COS 316', 'COS 316', 'COS 324', 'COS 324', 'COS 340', 'COS 418', 'COS 426', 'COS 445',
           'COS 445', 'COS 484', 'COS 488', 'COS 495', 'COS 511', 'COS 513', 'COS 518', 'COS 522', 'COS 585', 'COS 598D', 'EGR 154']


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
    return gen_name() if random.randint(0, 9) == 9 else ""


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


def gen_course_num():
    return gen_course_name()[4:]


def gen_match_list():
    return ', '.join([gen_course_name() for _ in range(random.choices([1, 2, 3, 4], weights=(4, 3, 2, 1), k=1)[0])])


def gen_course_assignment():
    return "" if _weighted_rand(0, 1, (1, 11)) else gen_course_num()


def gen_timestamp():
    return datetime.datetime.now().strftime('%m/%d/%Y %H:%M:%S')


def generate_student_info_sheet(entries=5):
    matrix = []
    for _ in range(entries):
        track, year = gen_track_year()
        line = [gen_name(), gen_name(), gen_nickname(),
                gen_name().lower(), "yes", track, year, gen_name(), gen_advisor2(), gen_bank_join_score(), gen_bank_join_score(), gen_course_assignment(), ""]
        matrix.append(line)
    out = [["Last", "First", "Nickname", "NetId", "Form", "Track",
            "Year", "Advisor", "Advisor2", "Bank", "Join", "Course", "Notes"]]
    matrix = sorted(matrix, key=lambda x: x[0])
    return out + matrix
