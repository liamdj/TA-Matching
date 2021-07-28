import csv
from os import replace
from typing import OrderedDict
import pandas as pd
import sys
import numpy as np
from collections import OrderedDict

if __name__ == "__main__":
   
    courses = pd.read_csv(sys.argv[1], index_col="Course", dtype = {"Course": str, "TA Slots": int}, usecols=["Course", "TA Slots"])
    courses["TA Slots"].fillna(1, inplace=True)

    slots = courses["TA Slots"].sum()

    course_slots = np.array([course for course, slots in courses.itertuples() for _ in range(slots)]).astype(str)

    with open("student_info.csv", "w", newline='') as info_file,  open("student_pref.csv", 'w', newline='') as pref_file:
        info_writer = csv.writer(info_file)
        info_writer.writerow(["Student", "Net ID", "Previous Courses", "Advisor's Course", "Assign Weight"])
        pref_writer = csv.writer(pref_file)
        pref_writer.writerow(["Username", "Name", "Excellent Matches", "Favorite Course", "Good Matches", "Okay Matches", "Poor Matches", "Unqualified Matches"])
        for i in range(slots):
            nums = [np.random.randint(2, 5), np.random.randint(2, 9), np.random.randint(0, 9), np.random.randint(0, 9)]
            np.random.shuffle(course_slots)
            ordering = list(OrderedDict.fromkeys(course_slots))
            excellent = ';'.join(ordering[0:nums[0]])
            good = ';'.join(ordering[nums[0]:nums[1]])
            okay = ';'.join(ordering[nums[1]:nums[2]])
            poor = ';'.join(ordering[nums[2]:nums[3]])
            unqualified = ';'.join(ordering[nums[3]:])
            favorite = np.random.choice(ordering[0:nums[0]])
            previous = ';'.join(np.random.choice(ordering[0:nums[1]], size=np.random.choice(3, p=[0.6, 0.3, 0.1]), replace=False))
            advisors = np.random.choice(ordering[0:nums[0]]) if np.random.random() < 0.33 else ''
            info_writer.writerow(["Student #{}".format(i), "student_{}".format(i), previous, advisors, ''])
            pref_writer.writerow(["student_{}@princeton.edu".format(i), "Student #{}".format(i), excellent, favorite, good, okay, poor, unqualified])

    students = ["Student #{}".format(i) for i in range(slots)]
    with open("prof_pref.csv", "w", newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Course","Excellent Matches","Good Matches","Okay Matches","Poor Matches","Unqualified Matches"])
        for course, slots in courses.itertuples():
            nums = [np.random.randint(slots + 1, 2 * slots + 1), np.random.randint(slots + 1, 2 * slots + 1), np.random.randint(0, 9), np.random.randint(0, 9)]
            np.random.shuffle(students)
            excellent = ';'.join(students[0:nums[0]])
            good = ';'.join(students[nums[0]:nums[1]])
            okay = ';'.join(students[nums[3]:])
            poor = ';'.join(students[nums[1]:nums[2]])
            unqualified = ';'.join(students[nums[2]:nums[3]])
            writer.writerow([course, excellent, good, okay, poor, unqualified])
