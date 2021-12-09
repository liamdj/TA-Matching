from ortools.graph import pywrapgraph
from scipy import stats
import numpy as np
import csv

DIGITS = 2


class MatchingGraph:

    def __init__(self, match_weights, student_weights, course_info, fixed_matches):

        self.num_students, self.num_courses, slots = len(student_weights), len(
            course_info.index), course_info['Slots'].sum()
        source, sink = range(self.num_students + self.num_courses +
                             slots, 2 + self.num_students + self.num_courses + slots)
        self.flow = pywrapgraph.SimpleMinCostFlow(sink + 1)

        # Each student cannot fill >1 course slot
        for i, w in enumerate(student_weights):
            self.flow.AddArcWithCapacityAndUnitCost(
                source, i, 1, -int(w * 10 ** DIGITS))

        # Each course slot cannot have >1 TA
        node = self.num_students + self.num_courses
        for i, (_, row) in enumerate(course_info.iterrows()):
            for s in range(int(row['Slots'])):
                self.flow.AddArcWithCapacityAndUnitCost(
                    self.num_students + i, node, 1, -MatchingGraph.fill_value(s, row['Base weight'], row['First weight']))
                self.flow.AddArcWithCapacityAndUnitCost(node, sink, 1, 0)
                node += 1

        # Force fixed matching edges to be filled
        missing = 0
        for _, row in fixed_matches.iterrows():
            si, ci = row["Student index"], row["Course index"]
            if si == -1:
                # self.flow.AddArcWithCapacityAndUnitCost(empty, students + ci, 1, 0)
                self.flow.SetNodeSupply(
                    self.num_students + ci, self.flow.Supply(self.num_students + ci) + 1)
            elif ci == -1:
                missing += 1
            else:
                match_cost = -int(match_weights[si, ci] * 10 ** DIGITS)
                # must include weight from incoming edge to student node
                assign_cost = -int(student_weights[si] * 10 ** DIGITS)
                self.flow.AddArcWithCapacityAndUnitCost(
                    si, self.num_students + ci, 1, match_cost + assign_cost)
                self.flow.SetNodeSupply(si, 1)

        # Edge weights given by preferences
        for si in range(self.num_students):
            if si not in fixed_matches["Student index"].values:
                for ci in range(self.num_courses):
                    if not np.isnan(match_weights[si, ci]):
                        cost = -int(match_weights[si, ci] * 10 ** DIGITS)
                        self.flow.AddArcWithCapacityAndUnitCost(
                            si, self.num_students + ci, 1, cost)

        # Attempt to fill max number of slots
        self.flow.SetNodeSupply(source, int(
            min(self.num_students, slots) - len(fixed_matches.index) + missing))
        self.flow.SetNodeSupply(sink, -int(min(self.num_students, slots)))

        # Option for not maximizing number of matches
        self.flow.AddArcWithCapacityAndUnitCost(
            source, sink, int(min(self.num_students, slots)), 0)

    # Value of filling slot is reciprocal with slot index

    def fill_value(index, base, first):
        return int((base + first / (index + 1)) * 10 ** DIGITS)

    def solve(self):
        return self.flow.Solve() == self.flow.OPTIMAL

    def get_matching(self, fixed_matches):
        matches = []
        for arc in range(self.flow.NumArcs()):
            si, ci = self.flow.Tail(arc), self.flow.Head(
                arc) - self.num_students
            # arcs from student to course
            if self.flow.Flow(arc) > 0 and si < self.num_students:
                matches.append((si, ci))
            # arcs from source to student
            elif ci < 0 and self.flow.Flow(arc) == 0:
                rows = fixed_matches.loc[fixed_matches['Student index'] == self.flow.Head(
                    arc)]
                if len(rows.index) == 0 or rows.iloc[0]['Course index'] == -1:
                    matches.append((self.flow.Head(arc), -1))
        return matches

    def write_matching(self, filename, weights, student_data, student_scores, course_data, course_scores, fixed_matches):

        matches = self.get_matching(fixed_matches)
        # put fixed matches at top and unassigned at bottom
        matches.sort(key=lambda tup: 100 if tup[1] == -1 else (-100 if (
            fixed_matches == tup).all(1).any() else -weights[tup[0], tup[1]]))

        student_perspective = stats.zscore(weights, axis=1, nan_policy='omit')
        course_perspective = stats.zscore(weights, axis=0, nan_policy='omit')

        # Note number of slots that were filled
        slots_filled = [0] * self.num_courses
        for _, ci in matches:
            if ci >= 0:
                slots_filled[ci] += 1

        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([
                "Netid", "Name", "Course", "Slots Filled", "Total Match Weight", "Fixed", "Previous", "Advisor-Advisee", "Negative Student Weight", "Student Rank", "Student Rank Score", "Matches Score (Student)", "Professor Rank", "Professor Rank Score", "Match Score (Course)"])

            for i, (si, ci) in enumerate(matches):
                student = student_scores.index[si]
                name = student_data.loc[student, "Name"]
                is_fixed = i < len(fixed_matches.index)
                is_previous = False
                is_advisor_advisee = False
                is_negative_student_weight = student_data.loc[student, "Weight"] < 0

                if ci >= 0:
                    course = course_scores.index[ci]
                    slots = course_data.loc[course, "Slots"]
                    s_rank = student_data.loc[student, course]
                    s_rank_score = student_scores.loc[student, course]
                    s_match_score = student_perspective[si, ci]
                    c_rank = course_data.loc[course, student]
                    c_rank_score = course_scores.loc[course, student]
                    c_match_score = course_perspective[si, ci]
                    is_previous = course in student_data.loc[student, "Previous"].split(
                        ';')
                    for instructor in course_data.loc[course, "Instructor"].split(';'):
                        is_advisor_advisee = is_advisor_advisee or instructor in student_data.loc[student, "Advisors"].split(
                        ';')
                    writer.writerow([
                        student, name, course, "{} / {}".format(slots_filled[ci], slots), "{:.2f}".format(weights[si, ci]), is_fixed, is_previous, is_advisor_advisee, is_negative_student_weight, s_rank, "{:.2f}".format(s_rank_score), "{:.2f}".format(s_match_score), c_rank, "{:.2f}".format(c_rank_score), "{:.2f}".format(c_match_score)])

                else:
                    writer.writerow(
                        [student, name, "unassigned", "", "", is_fixed, is_previous, is_advisor_advisee, is_negative_student_weight, "", "", "", "", "", ""])

    def print(self):
        for arc in range(self.flow.NumArcs()):
            print("start: {}, end: {}, capacity: {}, cost: {}, flow:{}".format(
                self.flow.Tail(arc), self.flow.Head(arc), self.flow.Capacity(arc), self.flow.UnitCost(arc), self.flow.Flow(arc)))
