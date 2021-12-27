import preprocess_sheets as preprocess
import matching
import write_to_google_sheets as write_gs
import g_sheet_consts as gs_consts


def preprocess_input_run_matching_and_write_matching(executor='UNCERTAIN',
                                                     input_dir_title='colab',
                                                     alternates=0,
                                                     planning_sheet_id=None,
                                                     student_preferences_sheet_id=None,
                                                     instructor_preferences_sheet_id=None):
    preprocess.write_csvs(
        output_directory_title=input_dir_title,
        planning_sheet_id=planning_sheet_id,
        student_prefs_sheet_id=student_preferences_sheet_id,
        instructor_prefs_sheet_id=instructor_preferences_sheet_id)

    num_executed = write_gs.get_num_execution_from_matchings_sheet()
    write_gs.copy_input_worksheets(
        num_executed, planning_sheet_id, student_preferences_sheet_id,
        instructor_preferences_sheet_id)

    run_and_write_matchings(
        executor, input_dir_title, num_executed, num_executed,
        alternates)


def run_and_write_matchings(executor, input_dir_title,
                            output_num_executed=None, input_num_executed=None,
                            alternates=0):
    """
    if `input_num_executed` is `None`, then use most recent copy
    """
    output_dir_path = f"data/{input_dir_title}"
    matching_weight, slots_unfilled, alt_weights = matching.run_matching(
        path=output_dir_path, alternates=alternates)
    print(f"unfilled slots: {slots_unfilled}")
    write_matchings(
        executor, output_dir_path, matching_weight, slots_unfilled,
        output_num_executed,
        input_num_executed, alt_weights)


def write_matchings(executor, output_dir_title, matching_weight, slots_unfilled=0,
                    output_num_executed=None, input_num_executed=None,
                    alt_weights=[]):
    """
    if `input_num_executed` is `None`, then use most recent copy
    """
    if input_num_executed is None:
        input_num_executed = str(write_gs.get_input_num_execution())
    if output_num_executed is None:
        output_num_executed = write_gs.get_num_execution_from_matchings_sheet(
            input_num_executed=input_num_executed)
    write_gs.write_output_csvs(
        len(alt_weights), output_num_executed, output_dir_title)
    write_gs.write_params_csv(output_num_executed, output_dir_title)
    write_gs.write_execution_to_ToC(
        executor, output_num_executed, matching_weight, slots_unfilled, alt_weights,
        input_num_executed=input_num_executed)
