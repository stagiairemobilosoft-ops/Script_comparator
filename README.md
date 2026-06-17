# Script_excel_comparator
Script Python avec Pandas pour comparer strictement deux fichiers Excel( ou csv) et valider le format des colonnes clients.
                     
# Refactorisation
--------------------------------------------------------------------------------------------------------
Dossier    |        Fichier                |           Classes           |          Fonctions          |
--------------------------------------------------------------------------------------------------------
constats   |       constant.py             |variable :STORE_ID_COLUMN / FILE_COLUMNS / SYNC_FIELDS /   |  
                                           |SCHEDULE_COLUMNS /CLIENT_COLUMN_RENAME / OPTIONAL_COLUMNS/ | 
                                           |RESULT_COLUMNS / STATUS_OPEN / STATUS_CLOSED / OUTPUT_DIR/ |
                                           |EXPORT_ACTION_MAP /OUTPUT_RESULTAT / OUTPUT_LOG /          |
                                           |DEFAULT_COMPARE_OUTPUT / DEFAULT_VALIDATE_OUTPUT /         |
                                           |ACTION_NO_CHANGE / ACTION_UPDATE / ACTION_CREATE /         |
                                           |ACTION_CLOSE / ACTION_COLUMN                               |                               
--------------------------------------------------------------------------------------------------------
normalize  |      file_loader.py           |-LoadError                  | - _is_client_format()        |
                                                                        | - _find_data_sheet()         | 
                                                                        | - _normalize_raw_dataframe() |
                                                                        | - load_store_file()          |
--------------------------------------------------------------------------------------------------------
comparison |          compare.py           |-ComparisonResult           |- _index_by_store()           | 
                                                                        |- _rows_to_dataframe()        | 
                                                                        |- compare_store_files()       |          
                                                                        |                              | 
            --------------------------------------------------------------------------------------------
           |          progress.py          |-ProgressReporter           |- update                      | 
                                                                        |- fiish                       |
            --------------------------------------------------------------------------------------------
           |        utils.py               |                            |- print-stage                 | 
                                                                        |- print_setp_done             |
                                                                        |- normalize_cell              | 
                                                                                                       |
--------------------------------------------------------------------------------------------------------
export     |     excel_export.py           |                            |- _cell_display_length()      | 
                                                                        |- autofit_columns()           |
                                                                        |- write_excel()               | 
                                                                        |- write_workbook()            |
                                                                        |- split_address()             | 
                                                                        |- _join_regular_hours()       |
                                                                        |- to_result_layout()          |
                                                                        |- build_sync_file()           |
                                                                        |- _sync_columns               |
                                                                        |- with_action                 |
            --------------------------------------------------------------------------------------------
                    log.py                 |                            |- build_analysis_summary()    | 
                                                                        |- write_analysis_workbook()   | 
                                                                        |- timestamp()                 | 
                                                                        |- timestamped_log_path()      |
                                                                        |- timestamped_analysis_path() |
                                                                        |- write_log()                 |
            --------------------------------------------------------------------------------------------
                    cmd.py                 |                            |- cmd_compare()               | 
                                                                        |- cmd_validate()              | 
                                                                        |- _ensure_output_dirs ()      | 
--------------------------------------------------------------------------------------------------------
validation |        quality_check.py       |                            |- _issue()                    | 
                                                                        |- _to_float                   | 
                                                                        |- deduplicate_stores          | 
                                                                        |- find_data_quality_issues()  | 
--------------------------------------------------------------------------------------------------------
          |       main.py                  |                            |- main()                      | 

# CMD
python3 main.py compare Nickel_Gmaps_20260211_.xlsx Nickel_Gmaps_20260522.xlsx
                           
