import params
import csv

new_scenarios = []
for s_pref in ["Favorite", "Good", "Okay"]:
    for c_pref in ["Favorite", "", "Veto"]:
        for prev_option in ["Previous", ""]:
            for advisors in ["2 Advisors", "1 Advisor", ""]:
                new_scenarios.append((s_pref, c_pref, prev_option, advisors))

with open('new_scenarios.csv', 'w+') as f:
    w = csv.writer(f)
    w.writerow(['Student Pref', 'Instructor Pref', 'Previous?', 'Advisor?'])
    w.writerows(new_scenarios)

combins = [("Favorite", "Favorite", params.FAVORITE_FAVORITE),
           ("Good", "Favorite", params.STUDENT_GOOD_INSTRUCTOR_FAVORITE),
           ("Okay", "Favorite", params.OKAY_COURSE_PENALTY),
           ("Favorite", "", params.STUDENT_FAVORITE_INSTRUCTOR_NEUTRAL),
           ("Good", "", 0), ("Okay", "", params.OKAY_COURSE_PENALTY)]
options = []
for s_pref, c_pref, w in combins:
    for prev_option, prev_w in [("Previous", params.PREVIOUS), ("", 0)]:
        for advisors, advisors_w in [("2 Advisors", params.ADVISORS * 2),
                                     ("1 Advisor", params.ADVISORS), ("", 0)]:
            options.append(
                (str(w + prev_w + advisors_w), s_pref, c_pref, prev_option,
                 advisors))

options.sort(key=lambda x: float(x[0]), reverse=True)

with open('cur_scenarios.csv', 'w+') as f:
    w = csv.writer(f)
    w.writerow(
        ['Weight', 'Student Pref', 'Instructor Pref', 'Previous?', 'Advisor?'])
    w.writerows(options)
