from typing import List, Tuple

import compare_outputs
import g_sheet_consts as gs_consts
import matching
import preprocess_sheets as preprocess
import write_to_google_sheets as write_gs


def preprocess_input_run_matching_and_write_matching(executor='UNCERTAIN',
                                                     input_dir_title='colab',
                                                     include_remove_and_add_ta=True,
                                                     include_remove_and_add_slot=True,
                                                     include_interviews=True,
                                                     alternates=0,
                                                     planning_sheet_id: str = None,
                                                     student_preferences_sheet_id: str = None,
                                                     instructor_preferences_sheet_id: str = None,
                                                     compare_matching_from_num_executed: str = None):
    num_executed = write_gs.get_num_execution_from_matchings_sheet()
    if compare_matching_from_num_executed is None:
        compare_matching_from_num_executed = f"{(int(num_executed) - 1):03d}"

    preprocess.write_csvs(
        previous_matching_ws_title=compare_matching_from_num_executed,
        output_directory_title=input_dir_title,
        planning_sheet_id=planning_sheet_id,
        student_prefs_sheet_id=student_preferences_sheet_id,
        instructor_prefs_sheet_id=instructor_preferences_sheet_id)

    input_copy_ids = write_gs.copy_input_worksheets(
        num_executed, planning_sheet_id, student_preferences_sheet_id,
        instructor_preferences_sheet_id)

    run_and_write_matchings(
        executor, input_dir_title, include_remove_and_add_ta,
        include_remove_and_add_slot, include_interviews, num_executed,
        num_executed, compare_matching_from_num_executed, alternates,
        input_copy_ids)


def run_and_write_matchings(executor: str, input_dir_title: str,
                            include_remove_and_add_ta=True,
                            include_remove_and_add_slot=True,
                            include_interviews=True,
                            output_num_executed: str = None,
                            input_num_executed: str = None,
                            compare_matching_from_num_executed: str = None,
                            alternates=0, input_copy_ids: Tuple[
            Tuple[str, str, str, str], Tuple[str, str], Tuple[
                str, str]] = None):
    """
    if `input_num_executed` is `None`, then use most recent copy
    """
    matching_weight, slots_unfilled, alt_weights, output_dir_path = run_matching(
        input_dir_title, alternates)
    write_matchings(
        executor, output_dir_path, matching_weight, include_remove_and_add_ta,
        include_remove_and_add_slot, include_interviews, slots_unfilled,
        output_num_executed, input_num_executed,
        compare_matching_from_num_executed, alt_weights, input_copy_ids)


def run_matching(input_dir_title: str, alternates=0) -> Tuple[
    float, int, List[float], str]:
    output_dir_path = f"data/{input_dir_title}"
    matching_weight, slots_unfilled, alt_weights = matching.run_matching(
        path=output_dir_path, alternates=alternates)
    if slots_unfilled > 0:
        print(f"\tunfilled slots: {slots_unfilled}")
    return matching_weight, slots_unfilled, alt_weights, output_dir_path


def write_matchings(executor: str, dir_path: str, matching_weight: float,
                    include_remove_and_add_ta=True,
                    include_remove_and_add_slot=True, include_interviews=True,
                    slots_unfilled=0, output_num_executed: str = None,
                    input_num_executed: str = None,
                    compare_matching_from_num_executed: str = None,
                    alt_weights: List[float] = [], input_copy_ids: Tuple[
            Tuple[str, str, str, str], Tuple[str, str], Tuple[
                str, str]] = None):
    """
    if `input_num_executed` is `None`, then use most recent copy;
    if `compare_matching_from_num_executed` is `None`, then do not compute diff
    """
    if input_num_executed is None:
        input_num_executed = str(write_gs.get_input_num_execution())
    if output_num_executed is None:
        output_num_executed = write_gs.get_num_execution_from_matchings_sheet(
            input_num_executed=input_num_executed)

    matching_output_sheet = write_gs.get_sheet(
        gs_consts.MATCHING_OUTPUT_SHEET_TITLE)

    outputs_dir_path = dir_path + '/outputs'
    include_matching_diff = False
    matching_diff_ws_title = None
    if compare_matching_from_num_executed:
        num_changes = compare_outputs.compare_and_write_matching_changes(
            outputs_dir_path, 'matchings.csv',
            dir_path + '/inputs/previous.csv')
        print(
            f"{num_changes} different assignments between {compare_matching_from_num_executed} and {output_num_executed}")
        if num_changes > 0:
            include_matching_diff = True
            matching_diff_ws_title = f'#{compare_matching_from_num_executed}->#{output_num_executed}'

    output_ids = write_gs.write_output_csvs(
        matching_output_sheet, include_remove_and_add_ta,
        include_remove_and_add_slot, len(alt_weights), output_num_executed,
        outputs_dir_path, matching_diff_ws_title)

    if not include_matching_diff and compare_matching_from_num_executed:
        matching_diff_ws_title = f'Same as #{compare_matching_from_num_executed}'

    param_copy_ids = write_gs.write_params_csv(
        output_num_executed, outputs_dir_path)
    toc_ws = write_gs.get_worksheet_from_sheet(
        matching_output_sheet, gs_consts.OUTPUT_TOC_TAB_TITLE)
    write_gs.write_execution_to_ToC(
        toc_ws, executor, output_num_executed, matching_weight, slots_unfilled,
        include_remove_and_add_ta, include_remove_and_add_slot,
        include_interviews, alt_weights, input_num_executed,
        matching_diff_ws_title, include_matching_diff, input_copy_ids,
        param_copy_ids, output_ids)
