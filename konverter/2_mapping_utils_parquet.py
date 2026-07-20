#%%
import pandas as pd
import pyarrow.parquet as pq
import pyarrow.dataset as ds
import duckdb

import os, sys, time, gc
from pathlib import Path
import re
import warnings
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
    category=UserWarning,
    module="openpyxl"
)
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
    "scenario": ["scenario", "Scenario", "Scenario name", "scenarioname", "Source Scenario", "scen", "SCEN1", "SCENARIO"],
    "region":   ["region", "Region", "Region name", "Source Region", "area", "AREA", "REGION", "Aggregate region"],
    "year":     ["year", "Year", "TIME", "Source Year", "Period", "YEAR"],
    "value":    ["value", "Value", "Source Value", "VAL", "growth", "VALUE", "IMPACT_VALUE"],
    "unit":     ["Unit", "unit", "UNIT", "IMPACT_UNIT"],
}

# ============================================================
# Hilfsklassen
# ============================================================

error_log = []
fatal_issues = []  # list of strings (global)

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def map_strict(df, column, mapping_dict, label, error_log, drop_unmapped=True, include_unit_in_missing=False):
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

    extra_cols = []
    if include_unit_in_missing and 'unit' in df.columns and column != 'unit':
        extra_cols.append('unit')
    # Only add unit to missing variables, not to unit itself
    # if 'unit' in df.columns and column != 'unit':
    #     extra_cols.append('unit')

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
    """
    Loads a mapping dictionary from an Excel sheet. 
    Used to load the various dictionary files from the global dictionary-Excel.
    This includes the dictionary for variables, models, scenarios.
    The dictionary for units is loaded with a different function, since it also includes conversion factors.
    The dictionary for regions is also now loaded with a different function.
    """
    df = pd.read_excel(file, sheet_name=sheet)
    if conv_col and conv_col not in df.columns:
        raise KeyError(f"Missing {conv_col} column in '{sheet}'.")
    if conv_col:
        mapping = {}
        for _, row in df.iterrows():
            src = row[src_col]
            tgt = row[tgt_col]
            factor = row[conv_col]
            if pd.notna(src) and pd.notna(tgt):
                mapping[src] = {'target': tgt, 'factor': factor if pd.notna(factor) else 1}
        return mapping
    else:
        return pd.Series(df[tgt_col].values, index=df[src_col]).to_dict()

def load_unit_pair_to_factor(file, sheet="units",
                             src_col="source_unit", tgt_col="target_unit", factor_col="conversion_factor"):
    df = pd.read_excel(file, sheet_name=sheet)

    for c in [src_col, tgt_col, factor_col]:
        if c not in df.columns:
            raise KeyError(f"Missing column '{c}' in sheet '{sheet}'")

    df = df[[src_col, tgt_col, factor_col]].copy()

    # normalize
    df[src_col] = df[src_col].map(lambda v: _norm_cell(v, empty_as_none=False))
    df[tgt_col] = df[tgt_col].map(lambda v: _norm_cell(v, empty_as_none=False))

    df[factor_col] = pd.to_numeric(df[factor_col], errors="coerce").fillna(1)

    # drop empty
    df = df[(df[src_col] != "") & (df[tgt_col] != "")].copy()

    # if duplicates exist, ensure they don't conflict
    dup = df.duplicated(subset=[src_col, tgt_col], keep=False)
    if dup.any():
        conflict = df[dup].groupby([src_col, tgt_col])[factor_col].nunique()
        bad = conflict[conflict > 1]
        if not bad.empty:
            raise ValueError(f"Conflicting conversion_factor for some (source_unit,target_unit) pairs in sheet '{sheet}'")

    # keep one row per pair
    return df.drop_duplicates(subset=[src_col, tgt_col]).set_index([src_col, tgt_col])[factor_col].to_dict()

def load_variable_target_units(file, sheet="variables", var_col="DE variable name", unit_col="unit"):
    """
    Return dict: {canonical_variable -> target_unit} from the variables sheet.
    canonical_variable must match df_input['variable'] AFTER your variable mapping.
    """
    df = pd.read_excel(file, sheet_name=sheet)

    for c in [var_col, unit_col]:
        if c not in df.columns:
            raise KeyError(f"Missing column '{c}' in sheet '{sheet}'")

    df = df[[var_col, unit_col]].copy()
    df[var_col] = df[var_col].astype("string").str.strip()
    df[unit_col] = df[unit_col].astype("string").str.strip()

    df = df[df[var_col].notna() & (df[var_col] != "") & df[unit_col].notna() & (df[unit_col] != "")]
    return df.set_index(var_col)[unit_col].to_dict()

def load_mapping_with_conflicts(file, sheet, src_col, tgt_col, *, extra_cols=None):
    """
    Loads mapping src->tgt but does NOT fail on duplicates.
    Instead returns:
      - mapping_unique: dict for src values that map uniquely to exactly one tgt
      - conflicts: dict[src] -> DataFrame(rows) for src values that map to multiple targets
      - raw_df: the normalized dataframe used (optional debugging)
    """
    df = pd.read_excel(file, sheet_name=sheet)

    for c in [src_col, tgt_col]:
        if c not in df.columns:
            raise KeyError(f"Missing column '{c}' in sheet '{sheet}'")

    cols = [src_col, tgt_col] + (extra_cols or [])
    df = df[cols].copy()

    df[src_col] = df[src_col].astype("string").str.strip()
    df[tgt_col] = df[tgt_col].astype("string").str.strip()

    # keep rows where src exists (target may be blank in your sheet; we keep it for reporting)
    df = df[df[src_col].notna() & (df[src_col] != "")].copy()

    # compute number of distinct targets per src (treat blanks as a value for conflict detection)
    # If you want blanks ignored in uniqueness, tell me; for now we treat blank as a target.
    n_targets = df.groupby(src_col)[tgt_col].nunique(dropna=False)

    conflicts_src = n_targets[n_targets > 1].index.tolist()

    conflicts = {}
    if conflicts_src:
        dups = df[df[src_col].isin(conflicts_src)].copy()
        for src in conflicts_src:
            conflicts[src] = dups[dups[src_col] == src].copy()

    # build unique mapping for src with exactly one target AND non-blank target
    unique_src = n_targets[n_targets == 1].index
    df_unique = df[df[src_col].isin(unique_src)].copy()

    # if there are blank targets among "unique", you probably don't want them as mapping entries
    df_unique_nonblank = df_unique[df_unique[tgt_col].notna() & (df_unique[tgt_col] != "")].copy()

    mapping_unique = pd.Series(df_unique_nonblank[tgt_col].values, index=df_unique_nonblank[src_col]).to_dict()

    return mapping_unique, conflicts, df

def load_region_mapping_model_aware(file, sheet="regions",
                                   src_col="source_region", tgt_col="target_region",
                                   folder_col="folder", model_sep="|"):
    """
    Builds two mappings from the regions sheet:
      - countries_map: {source_region -> target_region} for folder == 'countries' (must be unique)
      - regions_map_by_model: {(model, source_region) -> target_region} for folder == 'regions'
    Also returns:
      - conflicts: dict with human-readable conflict buckets for logging

    Assumptions (per your confirmation):
      - For folder=='regions', target_region is always 'MODEL_NAME|REGION_NAME'
      - Output should keep the MODEL_NAME| prefix
    """
    df = pd.read_excel(file, sheet_name=sheet)

    for c in [src_col, tgt_col, folder_col]:
        if c not in df.columns:
            raise KeyError(f"Missing column '{c}' in sheet '{sheet}'")

    df = df[[src_col, tgt_col, folder_col]].copy()
    df[src_col] = df[src_col].astype("string").str.strip()
    df[tgt_col] = df[tgt_col].astype("string").str.strip()
    df[folder_col] = df[folder_col].astype("string").str.strip().str.lower()

    # keep rows where source exists
    df = df[df[src_col].notna() & (df[src_col] != "")].copy()

    # split
    df_countries = df[df[folder_col].eq("countries")].copy()
    df_regions   = df[df[folder_col].eq("regions")].copy()

    conflicts = {
        "countries_ambiguous": {},   # src -> rows
        "regions_bad_format": {},    # src -> rows where target not like MODEL|...
        "regions_ambiguous": {},     # (model, src) -> rows if conflicting
    }

    # -----------------------
    # Countries: must be unique src -> tgt (ignoring blank targets still problematic)
    # -----------------------
    if not df_countries.empty:
        # consider blank target as value too (to catch 'GBR' -> blank + UK)
        n_targets = df_countries.groupby(src_col)[tgt_col].nunique(dropna=False)
        bad_src = n_targets[n_targets > 1].index.tolist()
        for src in bad_src:
            conflicts["countries_ambiguous"][src] = df_countries[df_countries[src_col] == src].copy()

        # build countries_map only for unique + non-blank targets
        ok_src = n_targets[n_targets == 1].index
        df_ok = df_countries[df_countries[src_col].isin(ok_src)].copy()
        df_ok = df_ok[df_ok[tgt_col].notna() & (df_ok[tgt_col] != "")]
        countries_map = pd.Series(df_ok[tgt_col].values, index=df_ok[src_col]).to_dict()
    else:
        countries_map = {}

    # -----------------------
    # Regions: require target format MODEL|Region
    # -----------------------
    if not df_regions.empty:
        has_sep = df_regions[tgt_col].fillna("").str.contains(r"\|", regex=True)
        df_bad = df_regions[~has_sep].copy()
        if not df_bad.empty:
            # group by src for logging
            for src, g in df_bad.groupby(src_col):
                conflicts["regions_bad_format"][src] = g.copy()

        df_good = df_regions[has_sep].copy()
        # parse model = part before first '|'
        df_good["__model"] = df_good[tgt_col].str.split(model_sep, n=1, expand=True)[0].astype("string").str.strip()

        # detect conflicting duplicates for same (model, src)
        key_cols = ["__model", src_col]
        # Here: conflicting if same key has multiple distinct target_region strings
        n_targets2 = df_good.groupby(key_cols)[tgt_col].nunique(dropna=False)
        bad_keys = n_targets2[n_targets2 > 1].index.tolist()
        for mdl, src in bad_keys:
            conflicts["regions_ambiguous"][(mdl, src)] = df_good[(df_good["__model"] == mdl) & (df_good[src_col] == src)].copy()

        # build mapping for keys with exactly 1 target
        ok_keys = n_targets2[n_targets2 == 1].index
        df_ok2 = df_good.set_index(key_cols)
        df_ok2 = df_ok2.loc[df_ok2.index.isin(ok_keys)].copy()
        regions_map_by_model = df_ok2[tgt_col].to_dict()
    else:
        regions_map_by_model = {}

    return countries_map, regions_map_by_model, conflicts, df

def _norm_cell(x, *, empty_as_none=True):
    """
    Normalize any cell-like value:
    - convert to string
    - replace NBSP with normal spaces
    - collapse repeated whitespace
    - strip
    """
    if pd.isna(x):
        return None if empty_as_none else ""

    s = str(x)
    s = s.replace("\u00A0", " ")       # NBSP -> normal space
    s = re.sub(r"\s+", " ", s)         # collapse whitespace runs
    s = s.strip()

    if s == "":
        return None if empty_as_none else ""
    return s

def norm_unit(x):
    return _norm_cell(x, empty_as_none=False)

def load_context_mapping(file, sheet, key_cols, value_cols):
    """
    Generic loader for context-aware mappings.
    Returns: dict[tuple] -> dict(value_cols...)
    Fails if the same key tuple appears with different values.
    """
    df = pd.read_excel(file, sheet_name=sheet)
    for c in key_cols + value_cols:
        if c not in df.columns:
            raise KeyError(f"Missing column '{c}' in sheet '{sheet}'")

    df = df[key_cols + value_cols].copy()
    for c in key_cols + value_cols:
        df[c] = df[c].apply(_norm_cell)

    # drop empty keys (if any key col is None, we still keep it as wildcard context)
    # but require that the "source" part is present: assume last key col is the "source" identifier
    src_key_col = key_cols[-1]
    df = df[df[src_key_col].notna()]

    mapping = {}
    collisions = []
    for _, row in df.iterrows():
        key = tuple(row[c] for c in key_cols)
        val = tuple(row[c] for c in value_cols)

        if key in mapping and mapping[key] != val:
            collisions.append((key, mapping[key], val))
        else:
            mapping[key] = val

    if collisions:
        lines = [f"[Dictionary] Conflicting rows in sheet '{sheet}' for the same key:"]
        for key, v_old, v_new in collisions[:50]:
            lines.append(f"- key={key}: {v_old} vs {v_new}")
        raise ValueError("\n".join(lines))

    # return as dict of dict for readability
    return {k: {value_cols[i]: v[i] for i in range(len(value_cols))} for k, v in mapping.items()}
     
def _next_copy_path(path: str, i: int) -> str:
    p = Path(path)
    return str(p.with_name(f"{p.stem}_copy{i}{p.suffix}"))

def log_error(msg, *, fatal=False):
    """Log to terminal and error_log; optionally also to fatal_issues."""
    # print: fatal in red, sonst gelb/cyan je nach Geschmack
    if fatal:
        print(Fore.RED + Style.BRIGHT + msg + Style.RESET_ALL)
        fatal_issues.append(msg)
    else:
        print(Fore.YELLOW + msg + Style.RESET_ALL)
    error_log.append(msg)

def parse_pipe_columns(spec: str):
    spec = (spec or '').strip()
    if not spec:
        return []
    return [c.strip() for c in spec.split('|')] if '|' in spec else [spec]

def resolve_columns_from_aliases(existing_cols, column_aliases, canonical_needed):
    """
    existing_cols: iterable of column names in the file (original)
    returns dict canonical -> actual_colname (only if found)
    """
    existing = set(existing_cols)
    resolved = {}
    for canonical in canonical_needed:
        for variant in column_aliases.get(canonical, []):
            if variant in existing:
                resolved[canonical] = variant
                break
    return resolved

def build_original_variable(df: pd.DataFrame, mapping_source_columns: str) -> pd.Series:
    cols = parse_pipe_columns(mapping_source_columns)
    if not cols:
        raise KeyError("Variable column is empty")

    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(f"Columns {missing} not found.")

    # cast+fill once per column (instead of inside the loop repeatedly)
    prepared = [df[c].astype('string').fillna('') for c in cols]

    s = prepared[0]
    for part in prepared[1:]:
        s = s + '|' + part

    return s.str.strip()

def ensure_numeric_value(series: pd.Series) -> pd.Series:
    # Fast path: already numeric (parquet)
    if series.dtype.kind in "if":  # int/float
        return series.astype("float64", copy=False)

    # Slow path: string cleanup (csv/excel)
    return pd.to_numeric(
        series.astype("string")
            .str.replace(" ", "", regex=False)
            .str.replace("\u00A0", "", regex=False)
            .str.replace(",", ".", regex=False),
        errors="coerce"
    )            

def q_ident(name: str) -> str:
    # DuckDB identifier quoting; double quotes inside name must be escaped
    return '"' + name.replace('"', '""') + '"'

def build_concat_expr(cols):
    # Cast to VARCHAR and coalesce NULL -> '' so concat doesn't become NULL
    parts = [f"COALESCE(CAST({q_ident(c)} AS VARCHAR), '')" for c in cols]
    return " || '|' || ".join(parts)

EXCEL_MAX_ROWS = 1_048_576
def pack_scenarios(rows_per_scenario: pd.Series, max_rows: int):
    """
    rows_per_scenario: Series index=scenario, values=rowcount (already sorted in desired order)
    Returns: list[list[scenario]] bins where sum(rows) <= max_rows
    Raises if a single scenario exceeds max_rows.
    """
    bins = []
    current = []
    current_rows = 0

    for scen, n in rows_per_scenario.items():
        n = int(n)
        if n > max_rows:
            raise ValueError(f"Scenario '{scen}' has {n:,} rows which exceeds Excel limit per file ({max_rows:,}).")

        if current and (current_rows + n > max_rows):
            bins.append(current)
            current = []
            current_rows = 0

        current.append(scen)
        current_rows += n

    if current:
        bins.append(current)

    return bins# ============================================================
# 1. Dictionary-Dateien laden
# ============================================================

print(f"Loading dictionary from: {DICTIONARY_FILE_PATH}")

dict_variable = load_mapping_dict(DICTIONARY_FILE_PATH, 'variables', 'names mapping', 'DE variable name')
dict_model    = load_mapping_dict(DICTIONARY_FILE_PATH, 'models', 'source_models', 'target_models')
dict_scenario = load_mapping_dict(DICTIONARY_FILE_PATH, 'scenarios', 'source_scenario', 'target_scenario')

# NEW: variable -> target unit (from variables-sheet)
var_to_target_unit = load_variable_target_units(DICTIONARY_FILE_PATH, 'variables', 'DE variable name', 'unit')
unit_pair_to_factor = load_unit_pair_to_factor(DICTIONARY_FILE_PATH, sheet="units")

# new - stricter loading of regions to catch duplicates and mapping issues early (since regions are often the main source of headaches in such mappings)
countries_map, regions_map_by_model, region_conflicts, df_region_dict = load_region_mapping_model_aware(
    DICTIONARY_FILE_PATH,
    sheet="regions",
    src_col="source_region",
    tgt_col="target_region",
    folder_col="folder",
    model_sep="|"
)

# --- NEW: report region dictionary conflicts early (and fail if you want)
if region_conflicts.get("countries_ambiguous"):
    log_error(f"[Dictionary] Countries region codes are ambiguous: {len(region_conflicts['countries_ambiguous'])}", fatal=False)
    for src, g in list(region_conflicts["countries_ambiguous"].items())[:50]:
        targets = sorted(set(g["target_region"].fillna("").astype(str).tolist()))
        log_error(f"  - {src}: {targets}", fatal=False)

if region_conflicts.get("regions_bad_format"):
    log_error(f"[Dictionary] Regions entries without MODEL| prefix: {len(region_conflicts['regions_bad_format'])}", fatal=False)
    for src, g in list(region_conflicts["regions_bad_format"].items())[:50]:
        targets = sorted(set(g["target_region"].fillna("").astype(str).tolist()))
        log_error(f"  - {src}: {targets}", fatal=False)

if region_conflicts.get("regions_ambiguous"):
    log_error(f"[Dictionary] Model-specific region mappings ambiguous: {len(region_conflicts['regions_ambiguous'])}", fatal=False)
    

print(f"{len(dict_variable)} variables loaded from dictionary.")
print(f"{len(dict_model)} models loaded from dictionary.")
print(f"{len(dict_scenario)} scenarios loaded from dictionary.\n")

print(f"{len(var_to_target_unit)} variable target units loaded from variables-sheet.")
print(f"{len(unit_pair_to_factor)} unit conversion pairs loaded from units-sheet.")

print(f"{len(countries_map)} country region mappings loaded (unique).")
print(f"{len(regions_map_by_model)} model-specific region mappings loaded (unique).")

# ============================================================
# 2. Mapping-Datei laden
# ============================================================

print(f"Reading dictionary file: {MAPPING_FILE_PATH}")
try:
    df_mapping_full = pd.read_excel(MAPPING_FILE_PATH, sheet_name='files').fillna('')
    FIRST_MAPPING_SHEET_NAME = pd.ExcelFile(MAPPING_FILE_PATH).sheet_names[0]
except FileNotFoundError:
    print(f"ERROR: Mapping-File '{MAPPING_FILE_PATH}' not found.")
    sys.exit(1)

# ============================================================
# 3. Gruppierung nach Quell-Dateien
# ============================================================

grouped_mappings = df_mapping_full.groupby(['File location', 'File name', 'Source model'])
print(f"\n{len(grouped_mappings)} unique files for processing found.")

# ============================================================
# current time for runtime measurement
# ============================================================
cur_time = time.time()
elapsed = cur_time - start_time
print(f"\n⏱️ Runtime so far: {elapsed:.2f} Seconds\n")

# ============================================================
# 4. Process all files grouped by model
# ============================================================

# Group only by model so all files of one model are collected together
model_groups = df_mapping_full.groupby('Source model')
print(f"\n{len(model_groups)} unique models for processing found.")

for model_raw, model_group in model_groups:
    model_key = dict_model.get(model_raw, model_raw)

    print(Fore.CYAN + Style.BRIGHT + f"\n=== Processing model: {model_key} (source: {model_raw}) ===" + Style.RESET_ALL)
    error_log.append(f"\n=== {model_key} (source: {model_raw}) ===")

    df_model_all = []  # collect IAMC data for each file of this model
    
    # --------------------------------------------------------
    # Loop through all files belonging to this model
    # --------------------------------------------------------
    for _, group_row in model_group.iterrows():
        file_location = group_row['File location']
        file_name     = group_row['File name']
        config        = group_row
        
        mapping_source_columns = str(config.get('Variable column', '')).strip()
        
        INPUT_FILE_PATH = os.path.join(MODEL_RESULTS_FOLDER, file_location, file_name)
        print(Fore.MAGENTA + Style.BRIGHT + f"\n--- File: {file_name} ---" + Style.RESET_ALL)
        error_log.append(f"\n--- {file_name} ---")

        # ----------------------------------------------------
        # Read source file (.xlsx, .csv or .parquet) into pandas DataFrame
        # ----------------------------------------------------
        sheet_name = config.get('Sheet name', 0) or 0
        t = time.time()

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
                
                
                # 1) header only
                cols = pd.read_csv(INPUT_FILE_PATH, sep=sep, nrows=0).columns.tolist()
                
                resolved = resolve_columns_from_aliases(cols, COLUMN_ALIASES, ['scenario','region','year','value','unit'])
                missing_base = [c for c in ['scenario','region','year','value','unit'] if c not in resolved]
                if missing_base:
                    raise KeyError(f"CSV missing required base columns (by alias): {missing_base}. Available: {cols}")

                base_needed = list(resolved.values())
                extra_needed = parse_pipe_columns(mapping_source_columns)

                usecols = list(dict.fromkeys(base_needed + extra_needed))

                # 2) full read but only needed cols
                df_input = pd.read_csv(INPUT_FILE_PATH, sep=sep, usecols=usecols, low_memory=False, engine="c")
            
            elif file_name.lower().endswith('.parquet'):
                # import pyarrow.dataset as ds

                # 1) Read schema (no data)
                parquet_cols = pq.ParquetFile(INPUT_FILE_PATH).schema.names

                # 2) Resolve base columns (scenario/region/year/value/unit) via aliases
                resolved = resolve_columns_from_aliases(
                    parquet_cols,
                    COLUMN_ALIASES,
                    canonical_needed=['scenario', 'region', 'year', 'value', 'unit']
                )
                missing_base = [c for c in ['scenario', 'region', 'year', 'value', 'unit'] if c not in resolved]
                if missing_base:
                    raise KeyError(
                        f"Parquet is missing required base columns (by alias): {missing_base}. "
                        f"Available columns: {parquet_cols}"
                    )

                base_needed = list(resolved.values())

                # 3) Add variable-source columns (exact from mapping excel)
                extra_needed = parse_pipe_columns(mapping_source_columns)

                # 4) Final list of columns to read
                needed_cols = list(dict.fromkeys(base_needed + extra_needed))
                
                # 5) prepare SQL statement           
                # use duckdb to concatenate the columns directly
                # make the SQL statement dynamic based on the columns we need to read
                parts = parse_pipe_columns(mapping_source_columns)  

                base_sql = ", ".join(q_ident(c) for c in base_needed)

                orig_expr = build_concat_expr(parts) + " AS original_variable"
                
                sql= f"""
                SELECT {base_sql}, {orig_expr}
                FROM read_parquet('{INPUT_FILE_PATH}')
                """
                
                con = duckdb.connect(database=":memory:")
                df_input = con.execute(sql).fetch_df()
                
                # 6) normalize string-like columns to pandas string dtype (consistent behavior)
                # only for parquet - maybe for others, too?
                for c in missing_base:
                    if c in df_input.columns:
                        df_input[c] = df_input[c].astype('string')
                
                # 6) Hard validation
                # missing_after = [c for c in needed_cols if c not in df_input.columns]
                missing_after = [c for c in base_needed if c not in df_input.columns]
                
                if missing_after:
                    raise KeyError(f"Parquet read returned missing columns: {missing_after}. Got: {list(df_input.columns)}")
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
        
        print(f"[TIMING] read took {time.time()-t:.2f}s")
        # ----------------------------------------------------
        # Also process files which are already in the IAMC format
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
        print("[TIMING] alias", time.time()-t); t=time.time()

        # ----------------------------------------------------
        # Variable column preparation
        # ----------------------------------------------------
        # bei Parquet-Dateien wird die Spalte 'original_variable' bereits beim Lesen erstellt, daher ist der folgende Schritt nur für andere Formate relevant.
        # df_input['original_variable'] = build_original_variable(df_input, mapping_source_columns)

        if 'original_variable' not in df_input.columns:
            try:
                df_input['original_variable'] = build_original_variable(df_input, mapping_source_columns)
            except KeyError as e:
                msg = f"ERROR: {e}. Skipping file {file_name}"
                print(msg)
                error_log.append(msg)
                continue
        else:        
            print("[DEBUG] original_variable already present; skipping build.")
            pass
        print("[TIMING] origvar", time.time()-t); t=time.time()

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

        # numeric value + apply factor (your existing parse logic)
        df_input['value'] = ensure_numeric_value(df_input['value']) 
        
        # ----------------------------------------------------
        # Dictionary mapping
        # ----------------------------------------------------
        # --- Report ambiguous region keys only if they appear in this input file
        # --- Region mapping (model-aware for folder=regions, global for countries)
        if 'region' in df_input.columns:
            src_region_series = df_input['region'].astype('string').str.strip()

            # Helper: normalize for case-insensitive matching, while keeping canonical spelling from dictionary
            def _norm_key(v) -> str:
                base = _norm_cell(v, empty_as_none=False)
                return (base or "").casefold()

            # Build canonical lookups for "already target" values
            # This enables e.g. 'UNITED KINGDOM' -> 'United Kingdom' if the target exists that way.
            country_target_canon = {}
            for tgt in countries_map.values():
                if pd.isna(tgt):
                    continue
                k = _norm_key(tgt)
                if k and k not in country_target_canon:
                    country_target_canon[k] = tgt

            region_target_canon = {}
            for tgt in regions_map_by_model.values():
                if pd.isna(tgt):
                    continue
                k = _norm_key(tgt)
                if k and k not in region_target_canon:
                    region_target_canon[k] = tgt

            # 1) pass-through if already in target form for THIS model (MODEL|...)
            # Since you want model prefix in output, we allow already-prefixed values.
            already_prefixed = src_region_series.fillna("").str.startswith(f"{model_key}|")

            # 2) allow-target fallback: if input already equals a target_region (case-insensitive), keep canonical spelling
            norm_src = src_region_series.map(_norm_key)
            mapped_country_target = norm_src.map(country_target_canon)
            mapped_region_target_any = norm_src.map(region_target_canon)

            # 3) map model-specific regions: (model, source_region)
            # build keys for vectorized-ish mapping
            keys = list(zip([model_key] * len(src_region_series), src_region_series.tolist()))
            mapped_model = pd.Series(keys, index=df_input.index).map(regions_map_by_model)

            # 4) map countries (model-independent)
            mapped_country = src_region_series.map(countries_map)

            # combine with precedence:
            # already_prefixed -> keep as is
            # else if matches a known target -> keep canonical target spelling
            # else model-specific mapping
            # else country mapping
            # else NaN (handled below)
            final_region = src_region_series.where(already_prefixed, pd.NA)
            final_region = final_region.fillna(mapped_country_target)
            final_region = final_region.fillna(mapped_region_target_any)
            final_region = final_region.fillna(mapped_model)
            final_region = final_region.fillna(mapped_country)

            # log missing / drop unmapped like map_strict does
            missing_mask = final_region.isna()
            if missing_mask.any():
                missing_vals = (
                    src_region_series.loc[missing_mask]
                    .astype("string")
                    .str.strip()
                )

                # remove empties
                missing_vals = missing_vals[missing_vals.notna() & (missing_vals != "")]

                amb_countries = set(region_conflicts.get("countries_ambiguous", {}).keys())
                bad_regions   = set(region_conflicts.get("regions_bad_format", {}).keys())

                missing_raw = src_region_series.loc[missing_mask].astype("string").str.strip()
                missing_raw = missing_raw[missing_raw.notna() & (missing_raw != "")]

                is_amb = missing_raw.isin(list(amb_countries))
                is_bad = missing_raw.isin(list(bad_regions))

                if is_amb.any():
                    vals = sorted(missing_raw[is_amb].unique().tolist())
                    log_error(f"[Dictionary] Ambiguous country codes encountered (fix dictionary): {vals}", fatal=False)

                if is_bad.any():
                    vals = sorted(missing_raw[is_bad].unique().tolist())
                    log_error(f"[Dictionary] Region codes have bad format in dictionary (expected MODEL|...): {vals}", fatal=False)
                                        
                counts = missing_vals.value_counts(dropna=False)

                msg_header = f"[Dictionary] {len(counts)} Regions entries not found in dictionary:"
                print(Fore.YELLOW + Style.BRIGHT + msg_header + Style.RESET_ALL)
                error_log.append(msg_header)

                # print each missing region once (with count if >1)
                for region_name, n in counts.items():
                    line = f"{region_name}" if n > 1 else str(region_name)
                    print(line)
                    error_log.append(line)

                # drop unmapped rows
                df_input = df_input.loc[~missing_mask].copy()
                final_region = final_region.loc[~missing_mask].copy()

            df_input['region'] = final_region
        else:
            msg = "[Dictionary] Column 'region' not found in input; cannot map Regions."
            print(Fore.YELLOW + msg + Style.RESET_ALL)
            error_log.append(msg)
            df_input['region'] = pd.Series(dtype='string')

        # build allow-target mapping        
        df_input['variable'] = map_strict(df_input, 'original_variable', dict_variable, 'Variables', error_log, include_unit_in_missing=True)        
        # after variable mapping: drop unmapped rows immediately
        df_input = df_input[df_input['variable'].notna()].copy()     
        
        df_input['scenario'] = map_strict(df_input, 'scenario', dict_scenario, 'Scenarios', error_log, include_unit_in_missing=False)
        # again: after scenario mapping: drop unmapped rows immediately
        df_input = df_input[df_input['scenario'].notna()].copy()   
        
        # Optional: log rows where value couldn't be parsed
        bad_value_rows = df_input.loc[df_input['value'].isna(), ['original_variable'] + (['unit'] if 'unit' in df_input.columns else [])].head(20)
        if not bad_value_rows.empty:
            msg = "[Check] WARNING: Some 'value' entries are non-numeric and were set to NaN (showing up to 20 rows)."
            print(Fore.YELLOW + msg + Style.RESET_ALL)
            error_log.append(msg)
            for _, r in bad_value_rows.iterrows():
                error_log.append(str(r.to_dict()))
        
        # target unit per row (from variables-sheet; df_input['variable'] must be canonical)
        df_input['desired_unit'] = df_input['variable'].map(var_to_target_unit)
        
        # warn if target unit missing for some variables
        missing_desired = df_input['desired_unit'].isna() | (df_input['desired_unit'].astype('string').str.strip() == '')
        has_var = df_input['variable'].notna() & (df_input['variable'].astype('string').str.strip() != '')

        missing_desired = missing_desired & has_var
        if missing_desired.any():
            print("[DEBUG] variable NaN count:", int(df_input['variable'].isna().sum()))
            print("[DEBUG] desired_unit NaN count:", int(df_input['desired_unit'].isna().sum()))
            miss_vars = df_input.loc[missing_desired, ['variable']].drop_duplicates().head(50)
            log_error(f"[Dictionary] WARNING: {int(missing_desired.sum())} rows have variables without target unit (showing up to 50 unique variables).", fatal=False)

            for v in miss_vars['variable'].tolist():
                log_error(f"  - {v}", fatal=False)


        cur_unit = df_input['unit'].map(norm_unit)
        des_unit = df_input['desired_unit'].map(norm_unit)

        has_des = des_unit.notna() & (des_unit != '')

        # pass-through when already correct
        same_unit = has_des & (cur_unit == des_unit)

        df_input['conversion_factor'] = 1.0

        need_conv = has_des & cur_unit.notna() & (cur_unit != '') & (cur_unit != des_unit)
        keys = list(zip(cur_unit[need_conv].tolist(), des_unit[need_conv].tolist()))
        factors = pd.Series(keys, index=df_input.index[need_conv]).map(unit_pair_to_factor)
        df_input.loc[need_conv, 'conversion_factor'] = factors

        # missing conversion rules -> log + drop (strict)
        miss_factor = need_conv & df_input['conversion_factor'].isna()

        if miss_factor.any():
            # Which unit-pairs are missing?
            pairs = pd.DataFrame({
                'source_unit': cur_unit[miss_factor].tolist(),
                'target_unit': des_unit[miss_factor].tolist(),
                'variable': df_input.loc[miss_factor, 'variable'].astype('string').str.strip().tolist(),
                'original_variable': df_input.loc[miss_factor, 'original_variable'].astype('string').str.strip().tolist() if 'original_variable' in df_input.columns else [''] * int(miss_factor.sum()),
                'region': df_input.loc[miss_factor, 'region'].astype('string').str.strip().tolist() if 'region' in df_input.columns else [''] * int(miss_factor.sum()),
                'scenario': df_input.loc[miss_factor, 'scenario'].astype('string').str.strip().tolist() if 'scenario' in df_input.columns else [''] * int(miss_factor.sum()),
            })

            # Summarize missing pairs
            pair_counts = pairs.groupby(['source_unit', 'target_unit']).size().reset_index(name='n').sort_values('n', ascending=False)

            msg_header = f"[Dictionary] {len(pair_counts)} Unit conversion pairs not found in units-sheet (showing up to 50):"
            print(Fore.YELLOW + Style.BRIGHT + msg_header + Style.RESET_ALL)
            error_log.append(msg_header)

            for _, r in pair_counts.head(50).iterrows():
                line = f"{r['source_unit']} -> {r['target_unit']} (x{int(r['n'])})"
                print(line)
                error_log.append(line)

                # For each missing pair: show which mapped variables are responsible (top 15)
                sub = pairs[(pairs['source_unit'] == r['source_unit']) & (pairs['target_unit'] == r['target_unit'])].copy()
                var_counts = sub['variable'].value_counts().head(15)

                print("  affected variables (top 15):")
                error_log.append("  affected variables (top 15):")
                for var_name, vn in var_counts.items():
                    vline = f"    - {var_name} (x{int(vn)})"
                    print(vline)
                    error_log.append(vline)

                # Optional: show a few example original variables to quickly spot bad mappings
                if 'original_variable' in sub.columns:
                    ex = sub[['original_variable']].dropna().drop_duplicates().head(5)['original_variable'].tolist()
                    if ex:
                        print("  examples original_variable:")
                        error_log.append("  examples original_variable:")
                        for e in ex:
                            eline = f"    - {e}"
                            print(eline)
                            error_log.append(eline)

            # IMPORTANT: keep your strict behavior (drop rows with missing conversion rule)
            df_input = df_input.loc[~miss_factor].copy()

            # refresh after dropping
            cur_unit = df_input['unit'].map(norm_unit)
            des_unit = df_input['desired_unit'].map(norm_unit)
            has_des = des_unit.notna() & (des_unit != '')

        # set final unit to desired_unit where available
        df_input.loc[has_des, 'unit'] = df_input.loc[has_des, 'desired_unit']
       
        # optional cleanup
        df_input.drop(columns=['desired_unit'], inplace=True, errors='ignore')


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

        df_iamc['model'] = model_key
        if model_raw not in dict_model:
            msg = f"WARNING: Source model '{model_raw}' not found in dictionary."
            print(msg)
            error_log.append(msg)

        df_model_all.append(df_iamc)
        del df_input, df_iamc; gc.collect()

    # --------------------------------------------------------
    # Combine and save one result per model
    # --------------------------------------------------------
    if not df_model_all:
        print(Fore.YELLOW + f"No valid files for model {model_key}, skipping." + Style.RESET_ALL)
        continue

    df_model_combined = pd.concat(df_model_all, ignore_index=True, copy=False)

    print("[TIMING] mapping+conv", time.time()-t); t=time.time()

    # --------------------------------------------------------
    # 4.x Pivot & save (save even with renamed duplicates)
    # --------------------------------------------------------
    
    try:
        final_out_file = None
        
        # --------------------------------------------------------
        # Pivot: always one row per file (full time series)
        # --------------------------------------------------------
        idx_cols = ['model', 'scenario', 'region', 'variable', 'unit', 'file_location', 'file_name']
        series_cols = ['model', 'scenario', 'region', 'variable', 'unit']

        df_output = (
            df_model_combined
            .pivot_table(
                index=idx_cols,
                columns='year',
                values='value',
                aggfunc='first'   # robust if same (idx,year) appears twice
            )
            .reset_index()
        )

        print(f"[DEBUG] After pivot: df_output rows={len(df_output):,} (one row per file/series)")

        # normalize column names to strings
        df_output.columns = [str(c) for c in df_output.columns]

        # year columns (after pivot)
        year_cols = [c for c in df_output.columns if c.isdigit()]

        # --------------------------------------------------------
        # Build time-series signature per row (for "identical series" detection)
        # --------------------------------------------------------
        # Use rounding to avoid float noise; keep <NA> stable
        sig = (
            df_output[year_cols]
            .astype('Float64')
            .round(12)
            .astype('string')
            .fillna('<NA>')
            .agg('|'.join, axis=1)
        )
        df_output['__sig'] = sig

        # --------------------------------------------------------
        # 1) Drop identical series within the same (model,scenario,region,variable,unit)
        #    Keep exactly one representative row per signature.
        # --------------------------------------------------------
        before = len(df_output)
        df_output = df_output.sort_values(series_cols + ['file_location', 'file_name']).copy()
        df_output = df_output.drop_duplicates(subset=series_cols + ['__sig'], keep='first')
        after = len(df_output)

        if before != after:
            msg = f"[Check] Dropped {before-after} file-rows with identical time series (kept one representative)."
            print(Fore.GREEN + msg + Style.RESET_ALL)
            error_log.append(msg)
            print(f"[DEBUG] Identical time series dropped: {before-after:,} (kept one representative per signature)")

        # --------------------------------------------------------
        # 2) Conflicts: series that still have >1 distinct signature
        #    => these are true duplicates you want to review
        # --------------------------------------------------------
        n_sigs = df_output.groupby(series_cols)['__sig'].transform('nunique')
        conflict_mask = n_sigs > 1

        # duplicates sheet = only conflicts (with file columns + full series)
        df_dups_output = None
        if conflict_mask.any():
            df_dups_output = df_output.loc[conflict_mask].drop(columns=['__sig']).copy()

            # rename region for conflicts to dup_<region>_<n>
            df_output.loc[conflict_mask, '__dup_i'] = (
                df_output.loc[conflict_mask]
                .groupby(series_cols)
                .cumcount()
                .add(1)
            )

            df_output.loc[conflict_mask, 'region'] = (
                'dup_' + df_output.loc[conflict_mask, 'region'].astype('string') + '_' +
                df_output.loc[conflict_mask, '__dup_i'].astype('Int64').astype('string')
            )

            df_output.drop(columns=['__dup_i'], inplace=True, errors='ignore')


        n_conflict_rows = int(conflict_mask.sum())
        n_conflict_series = int(df_output.loc[conflict_mask, series_cols].drop_duplicates().shape[0]) if n_conflict_rows else 0
        if not df_dups_output is None and not df_dups_output.empty:
            print("[DEBUG] There are duplicates; writing to separate sheet in output file.")        
        if (n_conflict_rows > 0):
            print(f"[DEBUG] True conflicts: {n_conflict_series:,} series, {n_conflict_rows:,} rows")

            
        # cleanup
        df_output.drop(columns=['__sig'], inplace=True, errors='ignore')

        # for i in range(0, 26):
        #     candidate = out_file if i == 0 else _next_copy_path(out_file, i)
        #     try:
        #         os.makedirs(os.path.dirname(out_file), exist_ok=True)
        #         # df_output.to_excel(candidate, index=False, sheet_name='pyam_data')
                
        #         # with pd.ExcelWriter(candidate, engine="openpyxl") as writer:
        #         #     df_output.to_excel(writer, index=False, sheet_name="data") # renamed from 'pyam_data' to 'data' to pass upload

        #         #     # write duplicates sheet only if it exists
        #         #     if df_dups_output is not None and not df_dups_output.empty:
        #         #         df_dups_output.to_excel(writer, index=False, sheet_name="duplicates")
        #         with pd.ExcelWriter(candidate, engine="xlsxwriter") as writer:
        #             df_output.to_excel(writer, index=False, sheet_name="data")
        #             if df_dups_output is not None and not df_dups_output.empty:
        #                 df_dups_output.to_excel(writer, index=False, sheet_name="duplicates")
        #         final_out_file = candidate
        #         if i > 0:
        #             msg = f"\n[Save] Output was open/locked; wrote to fallback file: {Path(candidate).name}"
        #             print(Fore.YELLOW + msg + Style.RESET_ALL)
        #             error_log.append(msg)
        #         break
        #     except PermissionError:
        #         continue
        
        # --------------------------------------------------------
        # It is possible that models create very large files (e.g. CITS)
        # Then, the Excel limit of 1,048,576 rows is exceeded. 
        # In that case, we need to split the output into multiple files.
        # --------------------------------------------------------
        # Split into multiple Excel files if Excel row limit is exceeded
        # (Option A: pack scenarios into as few files as possible)
        # --------------------------------------------------------
        ROW_MARGIN = 10_000  # safety buffer for header/edge cases
        max_rows_per_file = EXCEL_MAX_ROWS - ROW_MARGIN

        total_rows = len(df_output)

        # Build list of (suffix, dataframe_part)
        parts = []

        if total_rows <= EXCEL_MAX_ROWS:
            parts = [("", df_output)]
        else:
            # rows per scenario, sorted desc for best packing (fewer files)
            rows_per_scen = (
                df_output.groupby("scenario", dropna=False)
                .size()
                .sort_values(ascending=False)
            )

            scenario_bins = pack_scenarios(rows_per_scen, max_rows=max_rows_per_file)

            # suffix naming: scen01-05, scen06-10, ... based on scenario count in packing order
            scen_counter = 0
            for bin_scenarios in scenario_bins:
                start = scen_counter + 1
                scen_counter += len(bin_scenarios)
                end = scen_counter

                suffix = f"scen{start:02d}-{end:02d}"
                df_part = df_output[df_output["scenario"].isin(bin_scenarios)].copy()
                parts.append((suffix, df_part))

            msg = f"[Save] Output exceeds Excel row limit ({total_rows:,} rows). Writing {len(parts)} files grouped by scenario."
            print(Fore.YELLOW + msg + Style.RESET_ALL)
            error_log.append(msg)
        
        # If there are no true conflicts, drop source columns from the main data sheet (for every part)
        KEEP_SOURCE_COLS = True  # set False for final export after dedup
        if (not KEEP_SOURCE_COLS) and (not conflict_mask.any()):
            parts = [(suffix, df_part.drop(columns=['file_location','file_name'], errors='ignore'))
                    for suffix, df_part in parts]
                
        out_base = os.path.join(OUTPUT_FOLDER, f"pyam_{model_key}")

        for suffix, df_part in parts:
            out_file = out_base + ("" if suffix == "" else f"_{suffix}") + ".xlsx"

            final_out_file = None
            for i in range(0, 26):
                candidate = out_file if i == 0 else _next_copy_path(out_file, i)
                try:
                    os.makedirs(os.path.dirname(out_file), exist_ok=True)

                    with pd.ExcelWriter(candidate, engine="xlsxwriter") as writer:
                        df_part.to_excel(writer, index=False, sheet_name="data")

                        # duplicates sheet only if it exists (optional: you can also split it later)
                        if df_dups_output is not None and not df_dups_output.empty:
                            # For now: write duplicates ONLY when not splitting, or if you want you can split it similarly by scenario.
                            # Simplest behavior: only write duplicates when not split:
                            if suffix == "":
                                df_dups_output.to_excel(writer, index=False, sheet_name="duplicates")

                    final_out_file = candidate
                    if i > 0:
                        msg = f"[Save] Output was open/locked; wrote to fallback file: {Path(candidate).name}"
                        print(Fore.YELLOW + msg + Style.RESET_ALL)
                        error_log.append(msg)
                    break
                except PermissionError:
                    continue

            if final_out_file is None:
                raise PermissionError(f"Could not write output file (file locked?) after retries: {out_file}")

            print(Fore.GREEN + f"✅ Saved file for model {model_key}: {final_out_file}" + Style.RESET_ALL)

    except Exception as e:
        msg = f"ERROR during pivot/save for model {model_key}: {e}"
        print(Fore.RED + msg + Style.RESET_ALL)
        error_log.append(msg)
        continue

    # Clean up memory
    del df_output, df_model_all, df_model_combined
    gc.collect()

    cur_time = time.time()
    print(f"\n⏱️ Runtime so far: {cur_time - start_time:.2f} Seconds\n")
    
    print("[TIMING] pivot+save", time.time()-t)

# ============================================================
# 5. Finalization & Logs
# ============================================================
if fatal_issues:
    print(Fore.RED + Style.BRIGHT + "\n❌ Fatal dictionary issues detected. See error_log.txt for details." + Style.RESET_ALL)
    # optional: stop with non-zero exit (so CI / batch knows it's broken)
    raise SystemExit(2)

print(Fore.GREEN + Style.BRIGHT + "\n✅ All files processed." + Style.RESET_ALL)

with open(os.path.join(OUTPUT_FOLDER,'error_log.txt'), "w", encoding="utf-8") as f:
    for line in error_log:
        f.write(str(line) + "\n")

end_time = time.time()
elapsed = end_time - start_time
print(f"\n⏱️ Runtime of the script: {elapsed:.2f} Seconds\n")