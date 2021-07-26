from ortools.graph import pywrapgraph
from copy import copy

import csv

DIGITS = 3

class MatchingGraph:

    def __init__(self, match_weights, student_weights, course_weights, course_slots, fixed_matches):

        students, courses, slots = len(student_weights), len(course_weights), sum(course_slots)
        source, sink = range(students + courses + slots, 2 + students + courses + slots)
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
        for _, row in fixed_matches.iterrows():
            si, ci = row["Student index"], row["Course index"] 
            self.flow.SetNodeSupply(si, 1)
            match_cost = -int(match_weights[si, ci] * 10 ** DIGITS)
            # assign_cost = -int(student_weights[si] * 10 ** DIGITS)
            self.flow.AddArcWithCapacityAndUnitCost(si, students + ci, 1, match_cost)
            # insures no other edges from student will be added
            match_weights[si, :] = 0

        # Edge weights given by preferences
        for si in range(students):
            for ci in range(courses):
                if match_weights[si, ci] > 0:
                    cost = -int(match_weights[si, ci] * 10 ** DIGITS)
                    self.flow.AddArcWithCapacityAndUnitCost(si, students + ci, 1, cost)

        # Attempt to fill max number of slots
        self.flow.SetNodeSupply(source, min(students, slots) - len(fixed_matches.index))
        self.flow.SetNodeSupply(sink, -min(students, slots))

        # Option for not maximizing number of matches
        self.flow.AddArcWithCapacityAndUnitCost(source, sink, min(students, slots), 0)


    # Value of filling slot is reciprocal with slot index
    def fill_value(cap, index, weight):
        return int(weight / (index + 1) * 10 ** DIGITS)

    def solve(self):
        return self.flow.Solve() == self.flow.OPTIMAL

    def write(self, filename, students, student_ranks, student_scores, courses, course_ranks, course_scores, fixed_matches):

        matches = []
        for arc in range(self.flow.NumArcs()):
            si, ci = self.flow.Tail(arc), self.flow.Head(arc) - len(students)
            if self.flow.Flow(arc) > 0 and si < len(students):
                student = students[si]
                course = courses[ci]
                matches.append((student, course, -self.flow.UnitCost(arc) / 10 ** DIGITS))

        matches.sort(key=lambda tup: -tup[2])

        # Note which matches were fixed
        fixed = set()
        for _, row in fixed_matches.iterrows():
            i = next(i for i, (s, c, _) in enumerate(matches) if s == row["Student"] and c == row["Course"])
            fixed.add(i)

        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Student", "Course", "Total Match Weight", "Notes", "Student Score", "Student Rank", "Professor Score", "Professor Rank"])

            for i, (student, course, weight) in enumerate(matches):
                s_rank = student_ranks.loc[student, course]
                s_score = student_scores.loc[student, course]
                c_rank = course_ranks.loc[course, student]
                c_score = course_scores.loc[course, student]
                notes = ""
                if i in fixed:
                    notes += "Fixed"
                writer.writerow([student, course, weight, notes, "{:.2f}".format(s_score), s_rank,  "{:.2f}".format(c_score), c_rank])

    def print(self):
        for arc in range(self.flow.NumArcs()):
            print("start: {}, end: {}, capacity: {}, cost: {}".format(self.flow.Tail(arc), self.flow.Head(arc), self.flow.Capacity(arc), self.flow.UnitCost(arc)))