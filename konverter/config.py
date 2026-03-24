# depending on where you start the script, you might need to adjust the paths
# if you run "python konverter/2_mapping_utils.py", the paths should be relative to the root folder
# if you run it within "konverter", the paths should be relative to that folder, e.g.: 
# MODEL_RESULTS_FOLDER = r'..\\input\\POC_1.0'  # Pfad zu den Modelldateien

# MODEL_RESULTS_FOLDER = r'input\\POC_2.0'  # Pfad zu den Modelldateien

# relevant für 1_lookup_files
MODEL_RESULTS_FOLDER = r'..\\input\\Plastic POC 2.0_2025.10'  # Pfad zu den Modelldateien

# relevant für 2_mapping_utils
MAPPING_FILE_PATH = r'..\\overview_files.xlsx'
# MAPPING_FILE_PATH = 'overview_files_variables.xlsx'
DICTIONARY_FILE_PATH = r'..\\dictionary_dataexplorer_variables_translation-local.xlsm'
OUTPUT_FOLDER = r'..\\output'  # Ordner für Ausgabedateien

#relevant für 3_import_csv:
datei_pfad_csv = r'..\\input\\variable_info\\yaml_update.csv'