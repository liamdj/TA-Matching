import csv
from typing import Tuple, List

import numpy as np
import pandas as pd
from ortools.graph import pywrapgraph

DIGITS = 2


def add_to_slots_from_fixed_matches(course_info, fixed_matches):
    pre_filled_slots = {}  # key = course, value = num slots filled
    for _, row in fixed_matches.iterrows():
        course = row["Course"]
        if course:
            pre_filled_slots[course] = 1 + pre_filled_slots.get(course, 0)
    for course, filled_slots in pre_filled_slots.items():
        course_info.at[course, 'Slots'] = max(
            course_info.loc[course, 'Slots'], filled_slots)


def fill_value(index: int, base: float, first: float) -> int:
    """ Value of filling slot is reciprocal with slot index """
    return int((base + first / (index + 1)) * 10 ** DIGITS)


class MatchingGraph:

    def __init__(self, match_weights: np.ndarray, student_weights, course_info,
                 fixed_matches):
        add_to_slots_from_fixed_matches(course_info, fixed_matches)

        self.num_students = len(student_weights)
        self.num_courses = len(course_info.index)
        self.slots = course_info['Slots'].sum()
        source, sink = range(
            self.num_students + self.num_courses + self.slots,
            2 + self.num_students + self.num_courses + self.slots)
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
                    self.num_students + i, node, 1, -fill_value(
                        s, row['Base weight'], row['First weight']))
                self.flow.AddArcWithCapacityAndUnitCost(node, sink, 1, 0)
                node += 1

        # Force fixed matching edges to be filled
        missing = 0
        for _, row in fixed_matches.iterrows():
            si, ci = row["Student index"], row["Course index"]
            if si == -1:
                # self.flow.AddArcWithCapacityAndUnitCost(empty, students + ci, 1, 0)
                self.flow.SetNodeSupply(
                    self.num_students + ci, self.flow.Supply(
                        self.num_students + ci) + 1)
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
        self.flow.SetNodeSupply(
            source, int(
                min(self.num_students, self.slots) - len(
                    fixed_matches.index) + missing))
        self.flow.SetNodeSupply(sink, -int(min(self.num_students, self.slots)))

        # Option for not maximizing number of matches
        self.flow.AddArcWithCapacityAndUnitCost(
            source, sink, int(
                min(
                    self.num_students, self.slots)), 0)

    def solve(self):
        return self.flow.Solve() == self.flow.OPTIMAL

    def get_matching(self, fixed_matches: pd.DataFrame, weights: np.ndarray) -> \
            List[Tuple[int, int]]:
        """Returns list of (si, ci) matches"""
        matches = []
        fixed_matches = fixed_matches[['Student index', 'Course index']]
        for arc in range(self.flow.NumArcs()):
            si = self.flow.Tail(arc)
            ci = self.flow.Head(arc) - self.num_students
            # arcs from student to course
            if self.flow.Flow(arc) > 0 and si < self.num_students:
                matches.append((si, ci))
            # arcs from source to student
            elif ci < 0 and self.flow.Flow(arc) == 0:
                rows = fixed_matches.loc[
                    fixed_matches['Student index'] == self.flow.Head(arc)]
                if len(rows.index) == 0 or rows.iloc[0]['Course index'] == -1:
                    matches.append((self.flow.Head(arc), -1))

        # put fixed matches at top and unassigned at bottom
        matches.sort(
            key=lambda tup: 100 if tup[1] == -1 else (
                -100 if (fixed_matches == tup).all(1).any() else -weights[
                    tup[0], tup[1]]))
        return matches

    def get_slots_filled(self, matches: List[Tuple[int, int]]) -> List[int]:
        slots_filled = [0] * self.num_courses
        for _, ci in matches:
            if ci >= 0:
                slots_filled[ci] += 1
        return slots_filled

    def get_slots_unfilled(self, slots_filled: List[int]) -> int:
        return self.slots - sum(slots_filled)

    def write_matching(self, filename: str, weights: np.ndarray,
                       student_data: pd.DataFrame, course_data: pd.DataFrame,
                       fixed_matches: pd.DataFrame) -> Tuple[
        float, int, List[Tuple[int, int]]]:

        matches = self.get_matching(fixed_matches, weights)
        slots_filled = self.get_slots_filled(matches)

        with open(filename, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(
                ["NetID", "Name", "Notes", "Course", "Slots Filled",
                 "Total Match Weight", "Fixed", "Year", "Bank", "Join",
                 "Previous", "Advisor-Advisee", "Student Weight",
                 "Student Rank", "Professor Rank"])

            output = []
            for i, (si, ci) in enumerate(matches):
                student = student_data.index[si]
                name = student_data.loc[student, "Name"]
                year = student_data.loc[student, "Year"]
                bank = student_data.loc[student, "Bank"]
                bank = "" if np.isnan(bank) else bank
                join = student_data.loc[student, "Join"]
                join = "" if np.isnan(join) else join
                is_fixed = i < len(fixed_matches.index)
                is_previous = False
                is_advisor_advisee = False
                s_weight = student_data.loc[student, "Weight"]
                notes = student_data.loc[student, "Notes"]

                if ci >= 0:
                    course = course_data.index[ci]
                    slots = course_data.loc[course, "Slots"]
                    s_rank = student_data.loc[student, course]
                    c_rank = course_data.loc[course, student]
                    is_previous = course in student_data.loc[
                        student, "Previous"].split(';')
                    for instructor in course_data.loc[
                        course, "Instructor"].split(';'):
                        is_advisor_advisee = is_advisor_advisee or instructor in \
                                             student_data.loc[
                                                 student, "Advisors"].split(';')
                    output.append(
                        [student, name, notes, course,
                         "{} / {}".format(slots_filled[ci], slots),
                         "{:.2f}".format(weights[si, ci]), is_fixed, year, bank,
                         join, is_previous, is_advisor_advisee,
                         s_weight, s_rank, c_rank])

                else:
                    output.append(
                        [student, name, notes, "unassigned", "", "", is_fixed,
                         year, bank, join, is_previous, is_advisor_advisee,
                         s_weight, "", ""])
            output = sorted(output, key=lambda x: x[3])
            writer.writerows(output)
            optimal_cost = self.graph_weight()
            return optimal_cost, self.get_slots_unfilled(slots_filled), matches

    def graph_weight(self):
        """ Requires that the graph has been solved before """
        return -self.flow.OptimalCost() / (10 ** DIGITS)

    def print(self):
        for arc in range(self.flow.NumArcs()):
            print(
                "start: {}, end: {}, capacity: {}, cost: {}, flow:{}".format(
                    self.flow.Tail(arc), self.flow.Head(arc),
                    self.flow.Capacity(arc), self.flow.UnitCost(arc),
                    self.flow.Flow(arc)))
