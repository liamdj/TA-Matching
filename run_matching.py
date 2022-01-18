import compare_outputs
import preprocess_sheets as preprocess
import matching
import write_to_google_sheets as write_gs
import g_sheet_consts as gs_consts


def preprocess_input_run_matching_and_write_matching(executor='UNCERTAIN',
                                                     input_dir_title='colab',
                                                     include_removal_and_additional=True,
                                                     alternates=0,
                                                     planning_sheet_id=None,
                                                     student_preferences_sheet_id=None,
                                                     instructor_preferences_sheet_id=None,
                                                     compare_matching_from_num_executed=None):
    preprocess.write_csvs(
        output_directory_title=input_dir_title,
        planning_sheet_id=planning_sheet_id,
        student_prefs_sheet_id=student_preferences_sheet_id,
        instructor_prefs_sheet_id=instructor_preferences_sheet_id)

    num_executed = write_gs.get_num_execution_from_matchings_sheet()
    input_copy_ids = write_gs.copy_input_worksheets(
        num_executed, planning_sheet_id, student_preferences_sheet_id,
        instructor_preferences_sheet_id)

    if compare_matching_from_num_executed is None:
        compare_matching_from_num_executed = f"{(int(num_executed) - 1):03d}"

    run_and_write_matchings(
        executor, input_dir_title, include_removal_and_additional, num_executed,
        num_executed, compare_matching_from_num_executed, alternates,
        input_copy_ids)


def run_and_write_matchings(executor, input_dir_title,
                            include_removal_and_additional=True,
                            output_num_executed=None, input_num_executed=None,
                            compare_matching_from_num_executed=None,
                            alternates=0, input_copy_ids=None):
    """
    if `input_num_executed` is `None`, then use most recent copy
    """
    output_dir_path = f"data/{input_dir_title}"
    matching_weight, slots_unfilled, alt_weights = matching.run_matching(
        path=output_dir_path, alternates=alternates)
    if slots_unfilled > 0:
        print(f"\tunfilled slots: {slots_unfilled}")
    write_matchings(
        executor, output_dir_path, matching_weight,
        include_removal_and_additional, slots_unfilled, output_num_executed,
        input_num_executed, compare_matching_from_num_executed, alt_weights,
        input_copy_ids)


def write_matchings(executor, output_dir_title, matching_weight,
                    include_removal_and_additional=True, slots_unfilled=0,
                    output_num_executed=None, input_num_executed=None,
                    compare_matching_from_num_executed=None, alt_weights=[],
                    input_copy_ids=None):
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

    include_matching_diff = False
    matching_diff_ws_title = None
    if compare_matching_from_num_executed:
        student_changes, course_changes = compare_outputs.compare_matching_worksheet_with_csv(
            write_gs.get_worksheet_from_sheet(
                matching_output_sheet, compare_matching_from_num_executed),
            output_dir_title + "/outputs/matchings.csv")
        num_changes = compare_outputs.write_matchings_changes(
            student_changes, course_changes, output_dir_title)
        print(
            f"{num_changes} different assignments between {compare_matching_from_num_executed} and {output_num_executed}")
        if num_changes > 0:
            include_matching_diff = True
            matching_diff_ws_title = f'#{compare_matching_from_num_executed}->#{output_num_executed}'

    output_ids = write_gs.write_output_csvs(
        matching_output_sheet, include_removal_and_additional, len(alt_weights),
        output_num_executed, output_dir_title, matching_diff_ws_title)

    if not include_matching_diff and compare_matching_from_num_executed:
        matching_diff_ws_title = f'Same as #{compare_matching_from_num_executed}'

    param_copy_ids = write_gs.write_params_csv(
        output_num_executed, output_dir_title)
    toc_ws = write_gs.get_worksheet_from_sheet(
        matching_output_sheet, gs_consts.OUTPUT_TOC_TAB_TITLE)
    write_gs.write_execution_to_ToC(
        toc_ws, executor, output_num_executed, matching_weight, slots_unfilled,
        include_removal_and_additional, alt_weights, input_num_executed,
        matching_diff_ws_title, include_matching_diff, input_copy_ids,
        param_copy_ids, output_ids)
