from ortools.graph import pywrapgraph
from scipy import stats
import numpy as np
import csv

DIGITS = 2

class MatchingGraph:

    def __init__(self, match_weights, student_weights, course_weights, course_slots, fixed_matches):

        students, courses, slots = len(student_weights), len(course_weights), sum(course_slots)
        empty, source, sink = range(students + courses + slots, 3 + students + courses + slots)
        self.flow = pywrapgraph.SimpleMinCostFlow(sink + 1)

        # Each student cannot fill >1 course slot
        for i, w in enumerate(student_weights):
            self.flow.AddArcWithCapacityAndUnitCost(source, i, 1, -int(w * 10 ** DIGITS))

        # Each course slot cannot have >1 TA
        node = students + courses
        for i, (cap, value) in enumerate(zip(course_slots, course_weights)):
            for s in range(cap):
                self.flow.AddArcWithCapacityAndUnitCost(students + i, node, 1, -MatchingGraph.fill_value(cap, s, value))
                self.flow.AddArcWithCapacityAndUnitCost(node, sink, 1, 0)
                node += 1

        # Force fixed matching edges to be filled
        missing = 0
        for _, row in fixed_matches.iterrows():
            si, ci = row["Student index"], row["Course index"] 
            if si == -1:
                # self.flow.AddArcWithCapacityAndUnitCost(empty, students + ci, 1, 0)
                self.flow.SetNodeSupply(students + ci, self.flow.Supply(students + ci) + 1)
            elif ci == -1:
                missing += 1
            else:
                match_cost = -int(match_weights[si, ci] * 10 ** DIGITS)
                # must include weight from incoming edge to student node
                assign_cost = -int(student_weights[si] * 10 ** DIGITS)
                self.flow.AddArcWithCapacityAndUnitCost(si, students + ci, 1, match_cost + assign_cost)
                self.flow.SetNodeSupply(si, 1)

        # Edge weights given by preferences
        for si in range(students):
            if si not in fixed_matches["Student index"].values:
                for ci in range(courses):
                    if not np.isnan(match_weights[si, ci]):
                        cost = -int(match_weights[si, ci] * 10 ** DIGITS)
                        self.flow.AddArcWithCapacityAndUnitCost(si, students + ci, 1, cost)

        # Attempt to fill max number of slots
        self.flow.SetNodeSupply(source, min(students, slots) - len(fixed_matches.index) + missing)
        self.flow.SetNodeSupply(sink, -min(students, slots))

        # Option for not maximizing number of matches
        self.flow.AddArcWithCapacityAndUnitCost(source, sink, min(students, slots), 0)


    # Value of filling slot is reciprocal with slot index
    def fill_value(cap, index, weight):
        return int(weight / (index + 1) * 10 ** DIGITS)

    def solve(self):
        return self.flow.Solve() == self.flow.OPTIMAL

    def write_matching(self, filename, weights, student_data, student_scores, course_data, course_scores, fixed_matches):

        matches = []
        unassigned = []
        for arc in range(self.flow.NumArcs()):
            si, ci = self.flow.Tail(arc), self.flow.Head(arc) - len(student_scores.index)
            # arcs from student to course
            if self.flow.Flow(arc) > 0 and si < len(student_scores.index):
                matches.append((si, ci))
            # arcs from source to student
            elif ci < 0 and self.flow.Flow(arc) == 0 and self.flow.Head(arc) not in fixed_matches["Student index"].values:
                unassigned.append(self.flow.Head(arc))

        matches.sort(key=lambda tup: -weights[tup[0], tup[1]])
        student_perspective = stats.zscore(weights, axis=1, nan_policy='omit')
        course_perspective = stats.zscore(weights, axis=0, nan_policy='omit')

        # Note which matches were fixed
        fixed = set()
        for _, row in fixed_matches.iterrows():
            i = next(i for i, (s, c) in enumerate(matches) if s == row["Student index"] and c == row["Course index"])
            fixed.add(i)

        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Netid", "Name", "Course", "Total Match Weight", "Notes", "Student Rank", "Student Rank Score", "Matches Score (Student)", "Professor Rank", "Professor Rank Score", "Match Score (Course)"])

            for i, (si, ci) in enumerate(matches):
                student = student_scores.index[si]
                course = course_scores.index[ci]
                name = student_data.loc[student, "Name"]
                s_rank = student_data.loc[student, course]
                s_rank_score = student_scores.loc[student, course]
                s_match_score = student_perspective[si, ci]
                c_rank = course_data.loc[course, student]
                c_rank_score = course_scores.loc[course, student]
                c_match_score = course_perspective[si, ci]
                notes = []
                if i in fixed:
                    notes.append("fixed")
                if course in student_data.loc[student, "Previous"].split(';'):
                    notes.append("previous")
                if course == student_data.loc[student, "Advisors"]:
                    notes.append("advisor-advisee")
                writer.writerow([student, name, course, "{:.2f}".format(weights[si, ci]), ", ".join(notes), s_rank, "{:.2f}".format(s_rank_score), "{:.2f}".format(s_match_score), c_rank, "{:.2f}".format(c_rank_score), "{:.2f}".format(c_match_score)])

            for si in unassigned:
                writer.writerow([student_scores.index[si], "unassigned", "", "", "", "", "", "", "", ""])


    def print(self):
        for arc in range(self.flow.NumArcs()):
            print("start: {}, end: {}, capacity: {}, cost: {}, flow:{}".format(self.flow.Tail(arc), self.flow.Head(arc), self.flow.Capacity(arc), self.flow.UnitCost(arc), self.flow.Flow(arc)))