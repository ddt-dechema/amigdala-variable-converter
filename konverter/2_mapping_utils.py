#%%
import pandas as pd
import os
import sys
import time
from collections import defaultdict, Counter
from pathlib import Path
from colorama import Fore, Style, init
init(autoreset=True)

# ============================================================
# CONFIGURATION
# ============================================================

start_time = time.time()

# Zentrale Pfade
# MODEL_RESULTS_FOLDER = 'input\\POC_2.0_2025.10'  # Pfad zu den Modelldateien
MODEL_RESULTS_FOLDER = 'input\\POC_1.0'  # Pfad zu den Modelldateien
MAPPING_FILE_PATH = 'overview_files_variables_test.xlsx'
# MAPPING_FILE_PATH = 'overview_files_variables.xlsx'
DICTIONARY_FILE_PATH = 'dictionary_dataexplorer_variables_translation-local.xlsm'

# ============================================================
# 1. Dictionary-Dateien laden
# ============================================================

print(f"Lade Dictionary aus: {DICTIONARY_FILE_PATH}")

try:
    dict_var = pd.read_excel(DICTIONARY_FILE_PATH, sheet_name='variables')
    dict_reg = pd.read_excel(DICTIONARY_FILE_PATH, sheet_name='regions')
    dict_mod = pd.read_excel(DICTIONARY_FILE_PATH, sheet_name='models')
    dict_scen = pd.read_excel(DICTIONARY_FILE_PATH, sheet_name='scenarios')
    # (Optional für später) dict_units = pd.read_excel(DICTIONARY_FILE_PATH, sheet_name='units')
except Exception as e:
    raise RuntimeError(f"Fehler beim Laden der Dictionary-Datei: {e}")

# --- Variablen ---
if not {'DE variable name', 'names mapping'}.issubset(dict_var.columns):
    raise KeyError("Fehlende Spalten im Sheet 'variables': erwartet 'names mapping' und 'DE variable name'")
variable_dict = pd.Series(dict_var['DE variable name'].values,
                          index=dict_var['names mapping']).to_dict()

# --- Regionen ---
if not {'source_region', 'target_region'}.issubset(dict_reg.columns):
    raise KeyError("Fehlende Spalten im Sheet 'region': erwartet 'source_region' und 'target_region'")
region_dict = pd.Series(dict_reg['target_region'].values,
                        index=dict_reg['source_region']).to_dict()

# --- Modelle ---
if not {'source_models', 'target_models'}.issubset(dict_mod.columns):
    raise KeyError("Fehlende Spalten im Sheet 'models': erwartet 'source_models' und 'target_models'")
model_dict = pd.Series(dict_mod['target_models'].values,
                       index=dict_mod['source_models']).to_dict()

# --- Szenarien ---
if not {'source_scenario', 'target_scenario'}.issubset(dict_scen.columns):
    raise KeyError("Fehlende Spalten im Sheet 'scenarios': erwartet 'source_scenario' und 'target_scenario'")
scenario_dict = pd.Series(dict_scen['target_scenario'].values,
                         index=dict_scen['source_scenario']).to_dict()

print(f"{len(variable_dict)} variables loaded from dictionary.")
print(f"{len(region_dict)} regions loaded from dictionary.")
print(f"{len(model_dict)} models loaded from dictionary.")
print(f"{len(scenario_dict)} scenarios loaded from dictionary.\n")

# ============================================================
# 2. Initialisierung & Hilfsklassen
# ============================================================

error_log = []
unmapped_records = []

# --- Units & Normalisierung ---
try:
    import pint
    _UREG = pint.UnitRegistry()
except Exception:
    _UREG = None

# ============================================================
# 3. Mapping-Datei laden
# ============================================================

print(f"Reading dictionary file: {MAPPING_FILE_PATH}")
try:
    df_mapping_full = pd.read_excel(MAPPING_FILE_PATH, sheet_name='variable_mapping').fillna('')
    FIRST_MAPPING_SHEET_NAME = pd.ExcelFile(MAPPING_FILE_PATH).sheet_names[0]
except FileNotFoundError:
    print(f"ERROR: Mapping-File '{MAPPING_FILE_PATH}' not found.")
    sys.exit(1)

# ============================================================
# 4. Gruppierung nach Quell-Dateien
# ============================================================

grouped_mappings = df_mapping_full.groupby(['File location', 'File name', 'Source model'])
print(f"\n{len(grouped_mappings)} unique files for processing found.")

# ============================================================
# 5. Haupt-Schleife über Dateien
# ============================================================

for (file_location, file_name, model), group_of_mappings in grouped_mappings:
    config = group_of_mappings.iloc[0]
    # INPUT_FILE_PATH = os.path.join('input\\POC_2.0_2025.10', file_location, file_name)
    INPUT_FILE_PATH = os.path.join(MODEL_RESULTS_FOLDER, file_location, file_name)
    output_filename = f"pyam_{model}-{os.path.splitext(file_name)[0]}.xlsx"
    OUTPUT_FILE_PATH = os.path.join('output', output_filename)

    print(Fore.MAGENTA + Style.BRIGHT + f"\n--- Starting with process for: {file_name} ---" + Style.RESET_ALL)
    error_log.append(f"\n--- {file_name} ---")

    # --------------------------------------------------------------------
    # 5.1. Datei einlesen
    # --------------------------------------------------------------------
    sheet_name = config.get('Sheet name', 0) or 0
    try:
        if file_name.lower().endswith('.xlsx'):
            df_input = pd.read_excel(INPUT_FILE_PATH, sheet_name=sheet_name)
        elif file_name.lower().endswith('.csv'):
            sep = config['Separator'] if config['Separator'] else ','
            df_input = pd.read_csv(INPUT_FILE_PATH, sep=sep, low_memory=False)
        else:
            print("WARNING: Unknown Format – skipped.")
            error_log.append(f"WARNING: Unknown Format: {file_name}")
            continue
        print(f"File successfully loaded: {INPUT_FILE_PATH}")
    except Exception as e:
        print(f"ERROR reading file: {e}")
        error_log.append(f"ERROR reading file: {file_name} – {e}")
        continue

    # --------------------------------------------------------------------
    # 5.2. Variablen vorbereiten
    # --------------------------------------------------------------------
    def _to_clean_string(series: pd.Series) -> pd.Series:
        return series.astype('string').fillna('').str.strip()

    # --- Variable-Spalten (können eine oder mehrere sein) ---
    mapping_source_columns = str(config.get('Variable column', '')).strip()

    try:
        # Mehrfachspalten-Angabe? z. B. "Variable|Commodity"
        if '|' in mapping_source_columns:
            columns_to_combine = [col.strip() for col in mapping_source_columns.split('|')]

            # Existenz prüfen
            missing_cols = [c for c in columns_to_combine if c not in df_input.columns]
            if missing_cols:
                raise KeyError(f"Columns: {missing_cols} not found.")

            # Kombination zu einer zusammengesetzten Variablenbezeichnung
            cleaned = df_input[columns_to_combine].astype('string').fillna('').apply(lambda x: '|'.join(x), axis=1)
            df_input['original_variable'] = cleaned.str.strip()

        else:
            col = mapping_source_columns
            if col not in df_input.columns:
                raise KeyError(f"Column '{col}' not found.")
            df_input['original_variable'] = df_input[col].astype('string').fillna('').str.strip()

    except KeyError as e:
        print(f"ERROR: {e}. Skipped.")
        error_log.append(f"ERROR: {e}. Skipped. {file_name}")
        continue

    # Anschließend Variablen nach Dictionary übersetzen
    df_input['variable'] = df_input['original_variable'].map(variable_dict)

    # Fehlende Variablen erfassen
    missing_vars = df_input.loc[df_input['variable'].isna(), 'original_variable'].unique()
    if len(missing_vars) > 0:
        msg_header = f"[Dictionary] {len(missing_vars)} Variables not found in Dictionary:"
        print(Fore.RED + Style.BRIGHT + msg_header + Style.RESET_ALL)
        error_log.append(msg_header)
        for v in sorted(missing_vars):
            line = f"  - {v}"
            print(line)
            error_log.append(line)


    # ============================================================
    # Zusätzliche Prüfung: Regionen, Szenarien, Modelle, Variablen (je Zeile + farbig)
    # ============================================================

    def check_dictionary_entries(df, column, dictionary, label, error_log):
        """
        Prüft, ob Werte aus df[column] im dictionary vorkommen.
        Meldet fehlende Werte (Case- & Whitespace-insensitiv),
        listet sie zeilenweise im Log (copy-paste-freundlich) und färbt farbig ein.
        """
        if column not in df.columns:
            return

        # Werte bereinigen (String, Trim, Case)
        values = df[column].dropna().astype(str).str.strip()
        dict_keys = {k.strip().lower() for k in dictionary.keys() if isinstance(k, str)}
        missing = sorted({v for v in values if v.lower() not in dict_keys and v})

        if missing:
            msg_header = f"[Dictionary] {len(missing)} {label} not found in Dictionary:"
            print(Fore.YELLOW + Style.BRIGHT + msg_header + Style.RESET_ALL)
            error_log.append(msg_header)
            for val in missing:
                line = f"  - {val}"
                print(line)
                error_log.append(line)


    # --- Regionen prüfen ---
    if config.get('Source Region'):
        check_dictionary_entries(df_input, config['Source Region'], region_dict, 'Regionen', error_log)

    # --- Szenarien prüfen ---
    if config.get('Source Scenario'):
        check_dictionary_entries(df_input, config['Source Scenario'], scenario_dict, 'Szenarien', error_log)

    # --- Modelle prüfen ---
    if model not in model_dict:
        msg_header = "[Dictionary] Source model nicht im Dictionary gefunden:"
        print(Fore.YELLOW + Style.BRIGHT + msg_header + Style.RESET_ALL)
        model_line = f"  - {model}"
        print(model_line)
        error_log.append(msg_header)
        error_log.append(model_line)



    # alle Zeilen werden gelöscht, deren Variable nicht im Dictionary gefunden wurde
    df_input.dropna(subset=['variable'], inplace=True)
    if df_input.empty:
        print("INFO: No valid data after mapping. No output created.")
        continue

    # --------------------------------------------------------------------
    # 5.3. Transformation nach IAMC-Format
    # --------------------------------------------------------------------
    print("Transform files into IAMC-format ...")

    try:
        data_for_iamc = {
            'scenario': df_input[config['Source Scenario']],
            'region': df_input[config['Source Region']],
            'year': df_input[config['Source Year']],
            'value': df_input[config['Source Value']],
            'variable': df_input['variable']
        }
        df_iamc = pd.DataFrame(data_for_iamc)

        # Modelle
        df_iamc['model'] = str(model)
        if model in model_dict:
            df_iamc['model'] = model_dict[model]
        else:
            msg = f"WARNING: Source model '{model}' not found in dictionary."
            print(msg)
            error_log.append(msg)
           
        # Units - not yet implemented
        if config['Source Unit'] and config['Source Unit'] in df_input.columns:
            df_iamc['unit'] = df_input[config['Source Unit']]
        else:
            df_iamc['unit'] = 'undefined'

        # apply dictionary renaming for regions & Scenarios
        df_iamc['region'] = df_iamc['region'].map(region_dict).fillna(df_iamc['region'])
        df_iamc['scenario'] = df_iamc['scenario'].map(scenario_dict).fillna(df_iamc['scenario'])

    except KeyError as e:
        print(f"ERROR during transformation: {e}")
        error_log.append(f"ERROR during transformation: {file_name} – {e}")
        continue

    # --------------------------------------------------------------------
    # 5.4. Pivotieren & Speichern
    # --------------------------------------------------------------------
    print("Pivoting & Saving ...")

    try:
        df_output = (
            df_iamc.pivot_table(index=['model', 'scenario', 'region', 'variable', 'unit'],
                                columns='year', values='value', aggfunc='sum')
            .reset_index()
        )
        os.makedirs(os.path.dirname(OUTPUT_FILE_PATH), exist_ok=True)
        df_output.to_excel(OUTPUT_FILE_PATH, index=False, sheet_name='pyam_data')
        print(f"Result saved: {OUTPUT_FILE_PATH}")
    except Exception as e:
        print(f"ERROR during pivoting/saving: {e}")
        error_log.append(f"ERROR during pivoting: {file_name} – {e}")
        continue

# ============================================================
# 6. Abschluss & Logs
# ============================================================

print(Fore.GREEN + Style.BRIGHT + "\n✅ All files processed." + Style.RESET_ALL)

with open("output/error_log.txt", "w", encoding="utf-8") as f:
    for line in error_log:
        f.write(str(line) + "\n")

end_time = time.time()
elapsed = end_time - start_time
print(f"\n⏱️ Runtime of the script: {elapsed:.2f} Seconds\n")
