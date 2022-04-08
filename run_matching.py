import argparse
from typing import List, Tuple

import compare_outputs
import g_sheet_consts as gs_consts
import matching
import preprocess_sheets as preprocess
import write_to_google_sheets as write_gs


def preprocess_input_run_matching_and_write_matching(executor='UNCERTAIN',
                                                     input_dir_title='colab',
                                                     include_remove_and_add_features=True,
                                                     include_interviews=False,
                                                     alternates=0,
                                                     planning_sheet_id: str = None,
                                                     student_preferences_sheet_id: str = None,
                                                     instructor_preferences_sheet_id: str = None,
                                                     compare_matching_from_num_executed: str = None,
                                                     skip_previous=False):
    num_executed, matchings_sheet, matchings_worksheets, planning_input_copy_worksheets = write_gs.get_num_execution_from_matchings_sheet()
    if skip_previous:
        compare_matching_from_num_executed = None
    elif compare_matching_from_num_executed is None:
        compare_matching_from_num_executed = f"{(int(num_executed) - 1):03d}"

    _, _, _, planning_sheet_title, planning_sheet_worksheets = preprocess.write_csvs(
        previous_matching_ws_title=compare_matching_from_num_executed,
        output_directory_title=input_dir_title,
        planning_sheet_id=planning_sheet_id,
        student_prefs_sheet_id=student_preferences_sheet_id,
        instructor_prefs_sheet_id=instructor_preferences_sheet_id)

    input_copy_ids = write_gs.copy_input_worksheets(
        num_executed, planning_sheet_title, planning_sheet_worksheets,
        student_preferences_sheet_id, instructor_preferences_sheet_id)

    run_and_write_matchings(
        executor, input_dir_title, matchings_worksheets,
        include_remove_and_add_features, include_interviews, num_executed,
        num_executed, matchings_sheet, planning_input_copy_worksheets,
        compare_matching_from_num_executed, alternates, input_copy_ids)


def run_and_write_matchings(executor: str, input_dir_title: str,
                            matchings_worksheets: List[
                                write_gs.Worksheet] = None,
                            include_remove_and_add_features=True,
                            include_interviews=False,
                            output_num_executed: str = None,
                            input_num_executed: str = None,
                            matchings_sheet: write_gs.Spreadsheet = None,
                            planning_worksheets: List[
                                write_gs.Worksheet] = None,
                            compare_matching_from_num_executed: str = None,
                            alternates=0,
                            input_copy_ids: write_gs.InputCopyIDs = None):
    """
    if `input_num_executed` is `None`, then use most recent copy
    """
    matching_weight, slots_unfilled, alt_weights, output_dir_path = run_matching(
        input_dir_title, alternates, include_interviews)
    write_matchings(
        executor, output_dir_path, matching_weight, matchings_worksheets,
        matchings_sheet, planning_worksheets, include_remove_and_add_features,
        include_interviews, slots_unfilled, output_num_executed,
        input_num_executed, compare_matching_from_num_executed, alt_weights,
        input_copy_ids)


def run_matching(input_dir_title: str, alternates=0, run_interviews=False) -> \
        Tuple[float, int, List[float], str]:
    output_dir_path = f"data/{input_dir_title}"
    matching_weight, slots_unfilled, alt_weights = matching.run_matching(
        path=output_dir_path, alternates=alternates,
        run_interviews=run_interviews)
    if slots_unfilled > 0:
        print(f"\tunfilled slots: {slots_unfilled}")
    return matching_weight, slots_unfilled, alt_weights, output_dir_path


def write_matchings(executor: str, dir_path: str, matching_weight: float,
                    matchings_worksheets: List[write_gs.Worksheet],
                    matching_output_sheet: write_gs.Spreadsheet = None,
                    planning_copy_input_worksheets: List[
                        write_gs.Worksheet] = None,
                    include_remove_and_add_features=True,
                    include_interviews=False, slots_unfilled=0,
                    output_num_executed: str = None,
                    input_num_executed: str = None,
                    compare_matching_from_num_executed: str = None,
                    alt_weights: List[float] = [],
                    input_copy_ids: write_gs.InputCopyIDs = None):
    """
    if `input_num_executed` is `None`, then use most recent copy;
    if `compare_matching_from_num_executed` is `None`, then do not compute diff
    """
    if input_num_executed is None:
        input_num_executed, planning_copy_input_worksheets = write_gs.get_input_num_execution(
            planning_copy_input_worksheets)
        input_num_executed = str(input_num_executed)
    if output_num_executed is None:
        output_num_executed, matching_output_sheet, _, _ = write_gs.get_num_execution_from_matchings_sheet(
            input_num_executed=input_num_executed,
            matchings_sheet_worksheets=matchings_worksheets,
            planning_input_worksheets=planning_copy_input_worksheets)

    if matching_output_sheet is None:
        matching_output_sheet = write_gs.get_sheet(
            gs_consts.MATCHING_OUTPUT_SHEET_TITLE)

    outputs_dir_path = dir_path + '/outputs'
    include_matching_diff = False
    matching_diff_ws_title = None
    if compare_matching_from_num_executed:
        num_changes = compare_outputs.compare_and_write_matching_changes(
            outputs_dir_path, dir_path + '/inputs/previous.csv',
                              outputs_dir_path + '/matching.csv')
        print(
            f"{num_changes} different assignments between {compare_matching_from_num_executed} and {output_num_executed}")
        if num_changes > 0:
            include_matching_diff = True
            matching_diff_ws_title = f'#{compare_matching_from_num_executed}->#{output_num_executed}'

    output_ids = write_gs.write_output_csvs(
        matching_output_sheet, include_remove_and_add_features,
        include_interviews, len(alt_weights), output_num_executed,
        outputs_dir_path, matching_diff_ws_title)

    if not include_matching_diff and compare_matching_from_num_executed:
        matching_diff_ws_title = f'Same as #{compare_matching_from_num_executed}'

    param_copy_ids = write_gs.write_params_csv(
        output_num_executed, outputs_dir_path)
    toc_ws = write_gs.get_worksheet_from_sheet(
        matching_output_sheet, gs_consts.OUTPUT_TOC_TAB_TITLE)
    write_gs.write_execution_to_ToC(
        toc_ws, executor, output_num_executed, matching_weight, slots_unfilled,
        include_remove_and_add_features, include_interviews, alt_weights,
        input_num_executed, matching_diff_ws_title, include_matching_diff,
        input_copy_ids, param_copy_ids, output_ids)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'input_path', type=str,
        help='Name of the directory that contains the inputs directory')
    parser.add_argument(
        '--name', default='(nameless local machine)', type=str,
        help='Name of the person running the execution')
    parser.add_argument(
        '--alternates', default=2, metavar='alts', type=int, nargs='?',
        help='Number of alternate matchings to run with')
    parser.add_argument(
        '--compare_to', default=None, metavar='prev_comp', type=str, nargs='?',
        help='Tab name to compare against current run')
    parser.add_argument(
        '--exclude_add_remove', default=False, action='store_true',
        help='Write the removal and additional ta outputs')
    parser.add_argument(
        '--run_interviews', default=False,
        action='store_true', help='Run the interviews simulation')
    parser.add_argument(
        '--skip_previous', default=False,
        action='store_true', help='Do not boost from previous matching')
    parser.add_argument(
        '--planning', required=False, type=str,
        help='The Google Sheets id for the planning sheet')
    parser.add_argument(
        '--ta_prefs', required=False, type=str,
        help='The Google Sheets id for the student preferences sheet')
    parser.add_argument(
        '--fac_prefs', required=False, type=str,
        help='The Google Sheets id for the instructor preferences sheet')

    args = parser.parse_args()
    if args.planning and args.ta_prefs and args.fac_prefs:
        print("Preprocessing and running and writing")
        preprocess_input_run_matching_and_write_matching(
            args.name, args.input_path, not args.exclude_add_remove,
            args.run_interviews, args.alternates, args.planning, args.ta_prefs,
            args.fac_prefs, args.compare_to, args.skip_prvious)
    else:
        print("Running and writing")
        run_and_write_matchings(
            args.name, args.input_path,
            include_remove_and_add_features=not args.exclude_add_remove,
            include_interviews=args.run_interviews, alternates=args.alternates,
            compare_matching_from_num_executed=args.compare_to)
