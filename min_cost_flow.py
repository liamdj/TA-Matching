from ortools.graph import pywrapgraph
from scipy import stats
import csv

DIGITS = 2

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

    def write_matching(self, filename, weights, student_info, student_ranks, student_scores, courses, course_ranks, course_scores, fixed_matches):

        matches = []
        unassigned = []
        for arc in range(self.flow.NumArcs()):
            si, ci = self.flow.Tail(arc), self.flow.Head(arc) - len(student_info.index)
            # arcs from student to course
            if self.flow.Flow(arc) > 0 and si < len(student_info.index):
                matches.append((si, ci, -self.flow.UnitCost(arc) / 10 ** DIGITS))
            # arcs from source to student
            elif ci < 0 and self.flow.Flow(arc) == 0:
                unassigned.append(self.flow.Head(arc))

        matches.sort(key=lambda tup: -tup[2])
        student_perspective = stats.zscore(weights, axis=1)
        course_perspective = stats.zscore(weights, axis=0)

        # Note which matches were fixed
        fixed = set()
        for _, row in fixed_matches.iterrows():
            i = next(i for i, (s, c, _) in enumerate(matches) if s == row["Student"] and c == row["Course"])
            fixed.add(i)

        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Student", "Course", "Total Match Weight", "Notes", "Student Rank", "Student Rank Score", "Matches Score (Student)", "Professor Rank", "Professor Rank Score", "Match Score (Course)"])

            for i, (si, ci, weight) in enumerate(matches):
                student = student_info.index[si]
                course = courses[ci]
                s_rank = student_ranks.loc[student, course]
                s_rank_score = student_scores.loc[student, course]
                s_match_score = student_perspective[si, ci]
                c_rank = course_ranks.loc[course, student]
                c_rank_score = course_scores.loc[course, student]
                c_match_score = course_perspective[si, ci]
                notes = []
                if i in fixed:
                    notes.append("fixed")
                if course in student_info.loc[student, "Previous Courses"].split(';'):
                    notes.append("previous")
                if course == student_info.loc[student, "Advisor's Course"]:
                    notes.append("advisor-advisee")
                if course == student_info.loc[student, "Favorite Course"]:
                    notes.append("favorite")
                writer.writerow([student, course, weight, ", ".join(notes), s_rank, "{:.2f}".format(s_rank_score), "{:.2f}".format(s_match_score), c_rank, "{:.2f}".format(c_rank_score), "{:.2f}".format(c_match_score)])

            for si in unassigned:
                writer.writerow([student_info.index[si], "unassigned", "", "", "", "", "", "", "", ""])


    def print(self):
        for arc in range(self.flow.NumArcs()):
            print("start: {}, end: {}, capacity: {}, cost: {}".format(self.flow.Tail(arc), self.flow.Head(arc), self.flow.Capacity(arc), self.flow.UnitCost(arc)))