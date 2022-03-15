import csv
import math
import sys
from typing import List, Tuple, Dict

import numpy as np
import pandas as pd

import matching
import min_cost_flow

BucketsType = List[Tuple[float, float, float, List[Tuple[np.ndarray, float]]]]


def calculate_l2norm_for_uniform_entries(matched_students: int,
                                         denominator: int, ones: int) -> float:
    """
    Returns the L2 (Frobenius) norm for a matrix that would have `denominator`
    entries of `1/denominator` in each of `matched_students` rows, and `1` in
    `ones` rows.
    """
    return math.sqrt(matched_students * 1 / denominator + ones)


def add_noise_and_get_matching_interviews(stddev: float,
                                          base_weights: np.ndarray,
                                          student_weight_data: pd.DataFrame,
                                          course_graph_data: pd.DataFrame,
                                          fixed_matches: pd.DataFrame) -> Tuple[
    List[Tuple[int, int]], int]:
    noise = np.random.normal(0, stddev, base_weights.shape)
    new_weights = base_weights + noise
    graph = min_cost_flow.MatchingGraph(
        new_weights, student_weight_data, course_graph_data, fixed_matches)
    if not graph.solve():
        print('Problem optimizing flow')
        graph.print()
        return [], -1
    match = graph.get_matching(fixed_matches, base_weights)
    unfilled_slots = graph.get_slots_unfilled(graph.get_slots_filled(match))
    return match, unfilled_slots


def run_trials(sigma: float, trials_to_run: int,
               base_weights: np.ndarray,
               student_weight_data: pd.DataFrame,
               course_graph_data: pd.DataFrame,
               fixed_matches: pd.DataFrame) -> np.ndarray:
    courses = len(course_graph_data.index)
    students = len(student_weight_data.index)
    trial = np.zeros((students, courses), dtype=np.float64)
    for _ in range(trials_to_run):
        matching, unfilled = add_noise_and_get_matching_interviews(
            sigma, base_weights, student_weight_data, course_graph_data,
            fixed_matches)
        if unfilled == 0:
            for si, ci in matching:
                if si != -1 and ci != -1:
                    trial[si, ci] += 1.0
    return trial


def check_caps(course_data: pd.DataFrame, recent_simulation: np.ndarray,
               batch_num: int, num_trials: int,
               previous_simulations_percentages: np.ndarray) -> bool:
    courses = len(course_data.index)
    comparisons = np.absolute(
        (recent_simulation / num_trials) - previous_simulations_percentages)
    hard_cap = (comparisons >= 0.05).sum()
    if hard_cap == 0:
        soft_cap = (comparisons >= 0.02).sum(axis=0)
        courses_above_soft_cap = []
        for ci in range(courses):
            ci_slots = course_data.iloc[ci]['Slots']
            if soft_cap[ci] > int(ci_slots * 0.45):
                courses_above_soft_cap.append(
                    (course_data.index[ci], soft_cap[ci], ci_slots))

        if len(courses_above_soft_cap) == 0:
            print(
                f"Stopping after batch {batch_num} with {num_trials} trials due to being under the soft cap")
            return True
        else:
            print(
                f"After batch {batch_num} with {num_trials} trials, above soft cap ('Course', 'TAs > 2%', 'Slots'): {courses_above_soft_cap}")
    else:
        print(
            f"After batch {batch_num} with {num_trials} trials, still {hard_cap} differences above 0.05")
    return False


def insert_into_buckets(buckets: BucketsType,
                        simulation_percentages: np.ndarray, sigma: float):
    l2norm_to_insert = np.linalg.norm(simulation_percentages)
    for i, (low_thresh, max_sigma, min_sigma, sim_info) in enumerate(buckets):
        # use the fact that buckets is sorted from the smallest norm to largest
        # assume that any simulation will be less than the sigma=0 case
        if l2norm_to_insert < low_thresh:
            sim_info.append((simulation_percentages, sigma))
            new_min_sig = min(min_sigma, sigma)
            new_max_sig = max(max_sigma, sigma)
            buckets[i] = (low_thresh, new_max_sig, new_min_sig, sim_info)
            print(
                f"Inserting this simulation w/ norm {l2norm_to_insert} in bucket {i}")
            return


def run_single_simulation(sigma: float, weights: np.ndarray,
                          course_data: pd.DataFrame,
                          student_weight_data: pd.DataFrame,
                          course_graph_data: pd.DataFrame,
                          fixed_matches: pd.DataFrame):
    trials = 50
    sim_trials = trials
    sim_matches = run_trials(
        sigma, trials, weights, student_weight_data, course_graph_data,
        fixed_matches)
    for i in range(1, 10):
        trial_matches = run_trials(
            sigma, trials, weights, student_weight_data, course_graph_data,
            fixed_matches)
        sim_matches += trial_matches
        sim_trials += trials
        percent_decimals = trial_matches / trials
        if check_caps(
                course_data, sim_matches, i, sim_trials, percent_decimals):
            break
        trials *= 2
    simulation_percentages = sim_matches / sim_trials
    return simulation_percentages


def initialize_cumulative_percentages(students: int, courses: int,
                                      initial_matches: List[Tuple[
                                          int, int]]) -> np.ndarray:
    cumulative_percentages = np.zeros((students, courses), dtype=np.float64)
    for si, ci in initial_matches:
        if si != -1 and ci != -1:
            cumulative_percentages[si, ci] += 1.0
    return cumulative_percentages


def choose_sigma(buckets: BucketsType, bucket_index: int) -> float:
    def get_smallest_larger_min() -> float:
        for i in range(bucket_index - 1, -1, -1):
            if 0 < buckets[i][2] < sys.maxsize:
                return buckets[i][2]
        return 5.0

    def get_greatest_smaller_max() -> float:
        for i in range(bucket_index + 1, len(buckets)):
            if buckets[i][1] > 0:
                return buckets[i][1]
        return 0.0

    min_sigma = get_smallest_larger_min()
    max_sigma = get_greatest_smaller_max()

    return min_sigma * 0.5 + max_sigma * 0.5


def create_interview_list(course_data: pd.DataFrame, student_data: pd.DataFrame,
                          fixed_matches: pd.DataFrame, adjusted_path: str,
                          output_path: str,
                          initial_matches: List[Tuple[int, int]]):
    final_denominator = 4
    buckets = initialize_buckets(
        course_data[['Slots']].sum(),
        final_denominator, len(fixed_matches.index))

    # binary search one at a time until I get the desired length for a specific
    # bucket, but append any result to the buckets
    student_weight_data = student_data['Weight']
    course_graph_data = course_data[['Slots', 'Base weight', 'First weight']]
    cumulative_percentages = initialize_cumulative_percentages(
        len(student_data.index), len(course_data.index), initial_matches)

    weights = matching.match_weights(
        student_data, course_data, adjusted_path, 0.0, True)
    for simulation_num in range(len(buckets)):
        while len(buckets[simulation_num][3]) == 0:
            sigma = choose_sigma(buckets, simulation_num)
            print(
                f"Starting simulation for bucket {simulation_num} with sigma {sigma}")
            simulation_percentages = run_single_simulation(
                sigma, weights, course_data, student_weight_data,
                course_graph_data, fixed_matches)
            insert_into_buckets(buckets, simulation_percentages, sigma)

    buckets_to_print = []
    for desired_thresh, _, _, sim_list in buckets:
        average_sigma = 0.0
        for simulation, sigma in sim_list:
            cumulative_percentages += simulation / len(sim_list)
            average_sigma += sigma
        buckets_to_print.append((desired_thresh, average_sigma / len(sim_list)))
    print(f"Finished with buckets (thresholds, mean sigma): {buckets_to_print}")

    cumulative_percentages *= 100.0 / (len(buckets) + 1)
    weighted_percentages = parse_simulations_output(
        course_data, student_data, cumulative_percentages)
    write_simulations_output(output_path, weighted_percentages)


def initialize_buckets(filled_slots: int, final_denominator: int,
                       fixed_slots: int) -> BucketsType:
    initial_buckets = []  # sorted from highest to lowest
    for i in range(1, final_denominator + 1):
        initial_buckets.append(
            calculate_l2norm_for_uniform_entries(
                filled_slots - fixed_slots, i, fixed_slots))
    buckets = []
    for i, norm in enumerate(initial_buckets):
        if i < len(initial_buckets) - 1:
            next_norm = initial_buckets[i + 1]
            for j in np.arange(
                    norm, next_norm,
                    -(norm - next_norm) / (final_denominator - i)):
                buckets.append((float(j), -1.0, float(sys.maxsize), []))
        else:
            buckets.append((norm, -1.0, float(sys.maxsize), []))
    buckets.reverse()  # sort from the smallest norm to the largest
    print(f"Initializing buckets as {[round(b[0], 4) for b in buckets]}")
    return buckets


def parse_simulations_output(course_data: pd.DataFrame,
                             student_data: pd.DataFrame,
                             percentages: np.ndarray) -> Dict[
    str, List[Tuple[str, str, float]]]:
    readable_matches = {}
    courses = len(course_data.index)
    students = len(student_data.index)
    for ci in range(courses):
        sorted_matches = []
        for si in range(students):
            if percentages[si, ci] > 2.0:
                netid = student_data.index[si]
                name = student_data.iloc[si]['Name']
                percent = percentages[si, ci]
                sorted_matches.append((netid, name, percent))
        sorted_matches.sort(key=lambda x: x[2], reverse=True)
        readable_matches[course_data.index[ci]] = sorted_matches
    return readable_matches


def write_simulations_output(output_path: str, readable_matches: Dict[
    str, List[Tuple[str, str, float]]]):
    with open(output_path + 'interview_simulations.csv', 'w+') as f:
        writer = csv.writer(f)
        writer.writerow(['Course', 'NetID', 'Name', 'Percent Chance'])
        flattened = []
        for course, details in readable_matches.items():
            for netid, name, percentage in details:
                flattened.append([course, netid, name, round(percentage, 4)])
        writer.writerows(sorted(flattened, key=lambda x: x[0]))
