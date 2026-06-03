#%%
import pandas as pd
import os, sys, time, gc
import numpy as np
# from collections import defaultdict, Counter
from pathlib import Path

import errno

from colorama import Fore, Style, init
init(autoreset=True)

start_time = time.time()

# ============================================================
# CONFIGURATION & COLUMN ALIASES
# ============================================================

# Paths are now stored in config.py
from config import * 

COLUMN_ALIASES = {
    "scenario": ["scenario", "Scenario", "Scenario name", "Source Scenario", "scen", "SCEN2"],
    "region":   ["region", "Region", "Region name", "Source Region", "area", "AREA", "Aggregate region"],
    "year":     ["year", "Year", "TIME", "Source Year", "Period"],
    "value":    ["value", "Value", "Source Value", "VAL", "growth", "Net production", "Annual cost"],
    "unit":     ["Unit", "unit"],
}

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def map_strict(df, column, mapping_dict, label, error_log, drop_unmapped=True):
    """
    Maps a DataFrame column via a provided dictionary and logs missing mappings.
    Optionally drops unmapped rows for strict filtering.

    Parameters
    ----------
    df : pandas.DataFrame
        Input DataFrame
    column : str
        Column name in df to be mapped
    mapping_dict : dict
        Dictionary for mapping
    label : str
        Descriptive label for logging (e.g. 'Region', 'Scenario')
    error_log : list
        Global error log list
    drop_unmapped : bool, optional
        If True, removes rows with unmapped entries (default True)

    Returns
    -------
    pandas.Series
        The mapped series (NaNs removed if drop_unmapped=True)
    """
    if column not in df.columns:
        msg = f"[Dictionary] Column '{column}' not found in DataFrame for mapping {label}."
        print(Fore.YELLOW + msg + Style.RESET_ALL)
        error_log.append(msg)
        return pd.Series(dtype='string')

    mapped = df[column].map(mapping_dict)

    # find missing
    # missing_items = df.loc[mapped.isna(), column].unique().tolist()
    # also print unit of not found variables
    missing_rows = df.loc[mapped.isna(), [column] + ([ 'unit' ] if 'unit' in df.columns else [])].copy()

    # if missing_items:
    #     msg_header = f"[Dictionary] {len(missing_items)} {label} entries not found in dictionary:"
    #     print(Fore.YELLOW + Style.BRIGHT + msg_header + Style.RESET_ALL)
    #     error_log.append(msg_header)
    #     for val in sorted(missing_items):
    #         # line = f"  - {val}"
    #         line = f"{val}"
    #         print(line)
    #         error_log.append(line)

    extra_cols = []
    # Only add unit to missing variables, not to unit itself
    if 'unit' in df.columns and column != 'unit':
        extra_cols.append('unit')

    missing_rows = df.loc[mapped.isna(), [column] + extra_cols].copy()
    missing_rows = missing_rows.drop_duplicates()
    
    if not missing_rows.empty:
        msg_header = f"[Dictionary] {len(missing_rows)} {label} entries not found in dictionary:"
        print(Fore.YELLOW + Style.BRIGHT + msg_header + Style.RESET_ALL)
        error_log.append(msg_header)

        for _, row in missing_rows.drop_duplicates(subset=[column]).iterrows():
            val = row[column]
            if 'unit' in row and column != 'unit' and pd.notna(row['unit']):
                line = f"{val} - {row['unit']}"
            else:
                line = str(val)
            print(line)
            error_log.append(line)

        
    if drop_unmapped:
        df = df.loc[mapped.notna()].copy()
        mapped = mapped.dropna()

    return mapped

def load_mapping_dict(file, sheet, src_col, tgt_col, conv_col=None):
    # df = pd.read_excel(file, sheet_name=sheet, usecols=[src_col, tgt_col])
    df = pd.read_excel(file, sheet_name=sheet)
    if conv_col and conv_col not in df.columns:
        raise KeyError(f"Missing {conv_col} column in '{sheet}'.")
    if conv_col:
        # Liefert dict: {source_unit: {'target': ..., 'factor': ...}}
        mapping = {}
        for _, row in df.iterrows():
            src = row[src_col]
            tgt = row[tgt_col]
            factor = row[conv_col]
            if pd.notna(src) and pd.notna(tgt):
                mapping[src] = {'target': tgt, 'factor': factor if pd.notna(factor) else 1}
        return mapping
    else:
        # alter fallback
        return pd.Series(df[tgt_col].values, index=df[src_col]).to_dict()

        
def _next_copy_path(path: str, i: int) -> str:
    p = Path(path)
    return str(p.with_name(f"{p.stem}_copy{i}{p.suffix}"))
            
# ============================================================
# 1. Dictionary-Dateien laden
# ============================================================

print(f"Loading dictionary from: {DICTIONARY_FILE_PATH}")

dict_variable = load_mapping_dict(DICTIONARY_FILE_PATH, 'variables', 'names mapping', 'DE variable name')
dict_region   = load_mapping_dict(DICTIONARY_FILE_PATH, 'regions', 'source_region', 'target_region')
dict_model    = load_mapping_dict(DICTIONARY_FILE_PATH, 'models', 'source_models', 'target_models')
dict_scenario = load_mapping_dict(DICTIONARY_FILE_PATH, 'scenarios', 'source_scenario', 'target_scenario')
dict_unit     = load_mapping_dict(DICTIONARY_FILE_PATH, 'units', 'source_unit', 'target_unit', 'conversion_factor')
dict_unit_target = {k: v['target'] for k, v in dict_unit.items()}
dict_unit_factor = {k: v['factor'] for k, v in dict_unit.items()}


print(f"{len(dict_variable)} variables loaded from dictionary.")
print(f"{len(dict_region)} regions loaded from dictionary.")
print(f"{len(dict_model)} models loaded from dictionary.")
print(f"{len(dict_scenario)} scenarios loaded from dictionary.\n")
print(f"{len(dict_unit)} units loaded from dictionary.\n")

# ============================================================
# 2. Initialisierung & Hilfsklassen
# ============================================================

error_log = []

# ============================================================
# 3. Mapping-Datei laden
# ============================================================

print(f"Reading dictionary file: {MAPPING_FILE_PATH}")
try:
    df_mapping_full = pd.read_excel(MAPPING_FILE_PATH, sheet_name='files').fillna('')
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
# TO DO next steps:
# model_groups = df_mapping_full.groupby('Source model')


# ============================================================


# ============================================================
# current time for runtime measurement
# ============================================================
cur_time = time.time()
elapsed = cur_time - start_time
print(f"\n⏱️ Runtime so far: {elapsed:.2f} Seconds\n")

# ============================================================
# 5. Process all files grouped by model
# ============================================================

# Group only by model so all files of one model are collected together
model_groups = df_mapping_full.groupby('Source model')
print(f"\n{len(model_groups)} unique models for processing found.")

for model, model_group in model_groups:
    print(Fore.CYAN + Style.BRIGHT + f"\n=== Processing model: {model} ===" + Style.RESET_ALL)
    error_log.append(f"\n=== {model} ===")

    df_model_all = []  # collect IAMC data for each file of this model

    # --------------------------------------------------------
    # Loop through all files belonging to this model
    # --------------------------------------------------------
    for _, group_row in model_group.iterrows():
        file_location = group_row['File location']
        file_name     = group_row['File name']
        config        = group_row

        INPUT_FILE_PATH = os.path.join(MODEL_RESULTS_FOLDER, file_location, file_name)
        print(Fore.MAGENTA + Style.BRIGHT + f"\n--- File: {file_name} ---" + Style.RESET_ALL)
        error_log.append(f"\n--- {file_name} ---")

        # ----------------------------------------------------
        # Read source file (.xlsx or .csv)
        # ----------------------------------------------------
        sheet_name = config.get('Sheet name', 0) or 0
        try:
            if file_name.lower().endswith('.xlsx'):
                df_input = pd.read_excel(
                    INPUT_FILE_PATH,
                    sheet_name=sheet_name,
                    usecols=lambda col: col not in ["Unnamed: 0"],
                    engine="openpyxl"
                )
            elif file_name.lower().endswith('.csv'):
                sep = config['Separator'] if config['Separator'] else ','
                df_input = pd.read_csv(INPUT_FILE_PATH, sep=sep, low_memory=False, engine="c", dtype_backend="numpy_nullable")
                df_input.dropna(how='all', inplace=True)

            else:
                msg = f"WARNING: Unknown Format – skipped: {file_name}"
                print(msg)
                error_log.append(msg)
                continue
            print(f"File successfully loaded: {INPUT_FILE_PATH}")
        except Exception as e:
            msg = f"ERROR reading file {file_name}: {e}"
            print(msg)
            error_log.append(msg)
            continue

        # NEW - also process files which are already in the IAMC format
        # ----------------------------------------------------
        # Optional: Handle files that are already in pyam/IAMC format
        # ----------------------------------------------------
        pyam_like = all(
            any(str(year).isdigit() for year in df_input.columns)
            for _ in [0]
        ) and "variable" in df_input.columns and "region" in df_input.columns

        if pyam_like:
            print(Fore.CYAN + "Detected pyam/IAMC-wide format. Melting to long form..." + Style.RESET_ALL)

            # Melt Jahr-Spalten zu 'year'/'value'
            year_cols = [c for c in df_input.columns if str(c).isdigit()]
            df_input = df_input.melt(
                id_vars=[c for c in df_input.columns if c not in year_cols],
                value_vars=year_cols,
                var_name="year",
                value_name="value"
            )

            # Jahr-Spalte zu numerisch konvertieren
            df_input["year"] = pd.to_numeric(df_input["year"], errors="coerce")

            # Nur gültige Zeilen behalten
            df_input.dropna(subset=["year", "value"], inplace=True)


        # ----------------------------------------------------
        # 5.1.2  Standardize column names using aliases
        # ----------------------------------------------------
        for canonical, variants in COLUMN_ALIASES.items():
            for variant in variants:
                if variant in df_input.columns:
                    df_input.rename(columns={variant: canonical}, inplace=True)
                    break
        found_cols = [c for c in ["scenario", "region", "year", "value", "unit"] if c in df_input.columns]
        print(f"Standardized columns: {found_cols}")

        # ----------------------------------------------------
        # Variable column preparation
        # ----------------------------------------------------
        def _to_clean_string(series: pd.Series) -> pd.Series:
            return series.fillna('').astype('string', copy=False).str.strip()

        mapping_source_columns = str(config.get('Variable column', '')).strip()

        try:
            if '|' in mapping_source_columns:
                columns_to_combine = [col.strip() for col in mapping_source_columns.split('|')]
                missing_cols = [c for c in columns_to_combine if c not in df_input.columns]
                if missing_cols:
                    raise KeyError(f"Columns {missing_cols} not found.")
                cleaned = df_input[columns_to_combine].astype('string').fillna('').apply(lambda x: '|'.join(x), axis=1)
                df_input['original_variable'] = cleaned.str.strip()
                del cleaned; gc.collect()
            else:
                col = mapping_source_columns
                if col not in df_input.columns:
                    raise KeyError(f"Column '{col}' not found.")
                df_input['original_variable'] = df_input[col].astype('string').fillna('').str.strip()
        except KeyError as e:
            msg = f"ERROR: {e}. Skipping file {file_name}"
            print(msg)
            error_log.append(msg)
            continue

        # ----------------------------------------------------
        # Detect and handle IAMC/pyam wide-format files
        # ----------------------------------------------------
        year_cols = [c for c in df_input.columns if str(c).isdigit()]

        if year_cols:
            print(Fore.CYAN + f"Detected IAMC/pyam wide format with {len(year_cols)} year columns – converting to long format..." + Style.RESET_ALL)

            id_vars = [c for c in df_input.columns if c not in year_cols]
            df_input = df_input.melt(
                id_vars=id_vars,
                value_vars=year_cols,
                var_name="year",
                value_name="value"
            )

            # convert year -> numeric
            df_input["year"] = pd.to_numeric(df_input["year"], errors="coerce")

            # drop empty rows
            df_input.dropna(subset=["value"], inplace=True)
            df_input.reset_index(drop=True, inplace=True)


        # ----------------------------------------------------
        # Dictionary mapping
        # ----------------------------------------------------
        # build allow-target mapping
        _targets = set(dict_region.values())
        dict_region_allow_target = dict(dict_region)
        dict_region_allow_target.update({t: t for t in _targets if pd.notna(t)})
        
        df_input['variable'] = map_strict(df_input, 'original_variable', dict_variable, 'Variables', error_log)
        # df_input['region']   = map_strict(df_input, 'region', dict_region, 'Regions', error_log)
        df_input['region'] = map_strict(df_input, 'region', dict_region_allow_target, 'Regions', error_log)
        df_input['scenario'] = map_strict(df_input, 'scenario', dict_scenario, 'Scenarios', error_log)
        
        # --- Convert units into desired target unit/dimension
        # get conversion factor from dictionary (default to 1 if not found)
        df_input['conversion_factor'] = df_input['unit'].map(dict_unit_factor).fillna(1)

        # recalculate values based on conversion factor (if unit was found in dict, otherwise keep original value)
        # --- Ensure numeric values before applying conversion factor
        df_input['value'] = pd.to_numeric(
            df_input['value']
                .astype('string')
                .str.replace(' ', '', regex=False)     # remove thousands spaces
                .str.replace('\u00A0', '', regex=False) # remove NBSP if present
                .str.replace(',', '.', regex=False),   # decimal comma -> dot
            errors='coerce'
        )

        # Optional: log rows where value couldn't be parsed
        bad_value_rows = df_input.loc[df_input['value'].isna(), ['original_variable'] + (['unit'] if 'unit' in df_input.columns else [])].head(20)
        if not bad_value_rows.empty:
            msg = "[Check] WARNING: Some 'value' entries are non-numeric and were set to NaN (showing up to 20 rows)."
            print(Fore.YELLOW + msg + Style.RESET_ALL)
            error_log.append(msg)
            for _, r in bad_value_rows.iterrows():
                error_log.append(str(r.to_dict()))
        df_input['value'] = df_input['value'] * df_input['conversion_factor']

        # rename unit to target unit (if found in dict, otherwise keep original unit)
        df_input['unit'] = map_strict(df_input, 'unit', dict_unit_target, 'Units', error_log)


        df_input.dropna(subset=['variable', 'region', 'scenario'], inplace=True)
        if df_input.empty:
            msg = f"INFO: No valid data for {file_name}. Skipped."
            print(Fore.RED + msg + Style.RESET_ALL)
            error_log.append(msg)
            continue

        # ----------------------------------------------------
        # Transformation to IAMC format
        # ----------------------------------------------------
        print("Transforming to IAMC-format ...")
        data_for_iamc = {
            'scenario': df_input['scenario'],
            'region':   df_input['region'],
            'unit':     df_input['unit'],
            'year':     df_input['year'],
            'value':    df_input['value'],
            'variable': df_input['variable'],
            'file_location': file_location, # new, to faciliate debugging in case of duplicates
            'file_name': file_name  # new
        }
        df_iamc = pd.DataFrame(data_for_iamc)

        df_iamc['model'] = dict_model.get(model, model)
        if model not in dict_model:
            msg = f"WARNING: Source model '{model}' not found in dictionary."
            print(msg)
            error_log.append(msg)

        df_model_all.append(df_iamc)
        del df_input, df_iamc; gc.collect()

    # --------------------------------------------------------
    # Combine and save one result per model
    # --------------------------------------------------------
    if not df_model_all:
        print(Fore.YELLOW + f"No valid files for model {model}, skipping." + Style.RESET_ALL)
        continue

    df_model_combined = pd.concat(df_model_all, ignore_index=True, copy=False)

    # --------------------------------------------------------
    # Detect duplicates and mark them clearly
    # --------------------------------------------------------
    dup_cols = ['model', 'scenario', 'region', 'variable', 'unit', 'year']
    dupe_mask = df_model_combined.duplicated(subset=dup_cols, keep=False)

    if dupe_mask.any():
        dup_count = dupe_mask.sum()
        msg = f"[Check] Found {dup_count} duplicate rows for model {model}. Identical-valued duplicates will be removed; differing ones will be suffixed."
        print(Fore.YELLOW + msg + Style.RESET_ALL)
        error_log.append(msg)

        # identify duplicates grouped by keys
        grouped_dupes = df_model_combined[dupe_mask].groupby(dup_cols, dropna=False)

        rows_to_drop = set()
        rows_to_rename = []

        for key, group in grouped_dupes:
            # If all 'value' entries in group are identical, mark all but first for deletion
            if group['value'].nunique() == 1:
                rows_to_drop.update(group.index[1:])
            else:
                # assign incremental IDs for visible duplicates
                for i, idx in enumerate(group.index, start=1):
                    rows_to_rename.append((idx, f"dup_{group.iloc[i-1]['region']}_{i}"))
        # delete exact duplicates
        if rows_to_drop:
            df_model_combined.drop(index=list(rows_to_drop), inplace=True)
            msg = f"Removed {len(rows_to_drop)} rows with identical duplicates for model {model}."
            print(Fore.GREEN + msg + Style.RESET_ALL)
            error_log.append(msg)

        # rename only the true differing duplicates
        if rows_to_rename:
            for idx, new_name in rows_to_rename:
                df_model_combined.at[idx, 'region'] = new_name

            msg = f"Renamed {len(rows_to_rename)} remaining duplicate rows with 'dup_' prefix for model {model}."
            print(Fore.GREEN + msg + Style.RESET_ALL)
            error_log.append(msg)
    else:
        msg = f"[Check] No duplicates found for model {model}."
        print(msg)
        error_log.append(msg)

    # --------------------------------------------------------
    # 5.4. Pivotieren & Speichern (safe even with renamed duplicates)
    # --------------------------------------------------------
    try:
        final_out_file = None
        
        if dupe_mask.any():
            df_output = (
                df_model_combined
                .pivot(index=['model', 'scenario', 'region', 'variable', 'unit'
                            , 'file_location', 'file_name'], # new. to facilitate the detection of duplicates
                    columns='year', values='value')
                .reset_index()
            )
        else:
            df_output = (
                df_model_combined
                .pivot(index=['model', 'scenario', 'region', 'variable', 'unit'],
                    columns='year', values='value')
                .reset_index()
            )

        df_output.columns = [str(col) for col in df_output.columns]
        
        # --- Build duplicates sheet (only if duplicates exist) ---
        df_dups_output = None
        if dupe_mask.any():
            df_dups = df_model_combined.loc[dupe_mask].copy()

            # same pivot structure as the main output when duplicates exist
            df_dups_output = (
                df_dups
                .pivot(
                    index=['model', 'scenario', 'region', 'variable', 'unit',
                        'file_location', 'file_name'],
                    columns='year',
                    values='value'
                )
                .reset_index()
            )
            df_dups_output.columns = [str(col) for col in df_dups_output.columns]


        out_file = os.path.join(OUTPUT_FOLDER, f"pyam_{model}.xlsx")
        # os.makedirs(os.path.dirname(out_file), exist_ok=True)
        # df_output.to_excel(out_file, index=False, sheet_name='pyam_data')
        
        for i in range(0, 26):
            candidate = out_file if i == 0 else _next_copy_path(out_file, i)
            try:
                os.makedirs(os.path.dirname(out_file), exist_ok=True)
                # df_output.to_excel(candidate, index=False, sheet_name='pyam_data')
                
                with pd.ExcelWriter(candidate, engine="openpyxl") as writer:
                    df_output.to_excel(writer, index=False, sheet_name="pyam_data")

                    # write duplicates sheet only if it exists
                    if df_dups_output is not None and not df_dups_output.empty:
                        df_dups_output.to_excel(writer, index=False, sheet_name="duplicates")

                final_out_file = candidate
                if i > 0:
                    msg = f"\n[Save] Output was open/locked; wrote to fallback file: {Path(candidate).name}"
                    print(Fore.YELLOW + msg + Style.RESET_ALL)
                    error_log.append(msg)
                break
            except PermissionError:
                continue

        if final_out_file is None:
            raise PermissionError(f"Could not write output file (file locked?) after retries: {out_file}")

        output_msg_dups = "(with duplicates marked)" if dupe_mask.any() else ""
        print(Fore.GREEN + f"✅ Saved combined {output_msg_dups} file for model: {model} as {final_out_file}" + Style.RESET_ALL)

    except Exception as e:
        msg = f"ERROR during pivot/save for model {model}: {e}"
        print(Fore.RED + msg + Style.RESET_ALL)
        error_log.append(msg)
        continue

    # Clean up memory
    del df_output, df_model_all, df_model_combined
    gc.collect()

    cur_time = time.time()
    print(f"\n⏱️ Runtime so far: {cur_time - start_time:.2f} Seconds\n")
    
# ============================================================
# 6. Abschluss & Logs
# ============================================================

print(Fore.GREEN + Style.BRIGHT + "\n✅ All files processed." + Style.RESET_ALL)

with open(os.path.join(OUTPUT_FOLDER,'error_log.txt'), "w", encoding="utf-8") as f:
    for line in error_log:
        f.write(str(line) + "\n")

end_time = time.time()
elapsed = end_time - start_time
print(f"\n⏱️ Runtime of the script: {elapsed:.2f} Seconds\n")
