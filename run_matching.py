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
    sheet_ids = preprocess.write_csvs(output_directory_title=input_dir_title,
                                      planning_sheet_id=planning_sheet_id,
                                      student_prefs_sheet_id=student_preferences_sheet_id,
                                      instructor_prefs_sheet_id=instructor_preferences_sheet_id)

    run_and_write_matchings(executor, input_dir_title, *sheet_ids, alternates)


def run_and_write_matchings(executor, input_dir_title, planning_sheet_id=None,
                            student_preferences_sheet_id=None,
                            instructor_preferences_sheet_id=None, alternates=0):
    output_dir_path = f"data/{input_dir_title}"
    matching.run_matching(path=output_dir_path, alternates=alternates)
    write_matchings(executor, output_dir_path, planning_sheet_id,
                    student_preferences_sheet_id,
                    instructor_preferences_sheet_id, alternates)


def write_matchings(executor, output_dir_title, planning_sheet_id=None,
                    student_preferences_sheet_id=None,
                    instructor_preferences_sheet_id=None, alternates=0):
    num_executed = write_gs.get_num_execution_from_matchings_sheet(
        write_gs.get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE))
    write_gs.write_output_csvs(alternates, num_executed, output_dir_title)
    write_gs.write_execution_to_ToC(executor, num_executed, alternates,
                                    planning_sheet_id,
                                    student_preferences_sheet_id,
                                    instructor_preferences_sheet_id)

OFFICIAL_SPR22_TA_PLANNING_SHEET='1hOpAp7cdPyC1k018P0ANX7W6GFa9mALafeZwa4F9KkI'
OFFICIAL_SPR22_TA_PREFS_SHEET='102ScjAAywvAorVg4MGDzlHsQAd5PiRZTbT89nfNfwqU'
OFFICIAL_SPR22_INSTRUCTORS_SHEET='111z9ZiceHvrkWMV_zCxUQRfymF1l2cxAV1c8yVj2Vjw'
preprocess_input_run_matching_and_write_matching(
    executor='Nathan Local',
    alternates=0,
    planning_sheet_id=OFFICIAL_SPR22_TA_PLANNING_SHEET,
    student_preferences_sheet_id=OFFICIAL_SPR22_TA_PREFS_SHEET,
    instructor_preferences_sheet_id=OFFICIAL_SPR22_INSTRUCTORS_SHEET)
