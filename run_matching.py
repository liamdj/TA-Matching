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
    sheet_ids = preprocess.write_csvs(
        output_directory_title=input_dir_title,
        planning_sheet_id=planning_sheet_id,
        student_prefs_sheet_id=student_preferences_sheet_id,
        instructor_prefs_sheet_id=instructor_preferences_sheet_id)

    run_and_write_matchings(executor, input_dir_title, *sheet_ids, alternates)


def run_and_write_matchings(executor, input_dir_title, planning_sheet_id=None,
                            student_preferences_sheet_id=None,
                            instructor_preferences_sheet_id=None, alternates=0):
    output_dir_path = f"data/{input_dir_title}"
    matching_weight, alt_weights = matching.run_matching(
        path=output_dir_path, alternates=alternates)
    write_matchings(
        executor, output_dir_path, matching_weight, planning_sheet_id,
        student_preferences_sheet_id, instructor_preferences_sheet_id,
        alt_weights)


def write_matchings(executor, output_dir_title, matching_weight,
                    planning_sheet_id=None, student_preferences_sheet_id=None,
                    instructor_preferences_sheet_id=None, alt_weights=[]):
    num_executed = write_gs.get_num_execution_from_matchings_sheet(
        write_gs.get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE))
    write_gs.write_output_csvs(len(alt_weights), num_executed, output_dir_title)
    write_gs.write_execution_to_ToC(
        executor, num_executed, matching_weight, alt_weights, planning_sheet_id,
        student_preferences_sheet_id, instructor_preferences_sheet_id)
