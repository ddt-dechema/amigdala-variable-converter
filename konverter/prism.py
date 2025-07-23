import pandas as pd
import os

mapping_path = "variable_mapping.xlsx"
mapping = pd.read_excel(mapping_path)

input_base = "input/"
all_pyam_dfs = []
loaded_files = {}

for idx, row in mapping.iterrows():
    file_location = row['File location'] if pd.notna(row['File location']) else ''
    file_name = row['File name']
    source_model = row['Source model']
    variable_col = row['Variable column']  # z.B. 'sortingstream'
    orig_var_value = row['Variable value (original)']  # z.B. 'Recycling'
    pyam_var = row['Variable name (new)'] if pd.notna(row['Variable name (new)']) and row['Variable name (new)'] else orig_var_value
    unit = row['Unit'] if pd.notna(row['Unit']) else ''
    definition = row['Definition'] if 'Definition' in row and pd.notna(row['Definition']) else ''
    used_by = row['used by'] if 'used by' in row and pd.notna(row['used by']) else ''

    # Datei laden (Caching)
    full_path = os.path.join(input_base, file_location, file_name)
    if not os.path.isfile(full_path):
        print(f"Datei nicht gefunden: {full_path}")
        continue

    if full_path in loaded_files:
        df = loaded_files[full_path]
    else:
        if full_path.lower().endswith('.csv'):
            df = pd.read_csv(full_path)
        elif full_path.lower().endswith('.xlsx') or full_path.lower().endswith('.xls'):
            df = pd.read_excel(full_path)
        else:
            print(f"Nicht unterstütztes Dateiformat: {full_path}")
            continue
        loaded_files[full_path] = df

    # Prüfen, ob die Variable-Spalte existiert
    if variable_col not in df.columns:
        print(f"Variable-Spalte '{variable_col}' nicht gefunden in {file_name}")
        continue

    # Zeilen filtern, wo die Variable-Spalte den gewünschten Wert hat
    df_var = df[df[variable_col] == orig_var_value]
    if df_var.empty:
        print(f"Keine Zeilen mit {variable_col} == '{orig_var_value}' in {file_name}")
        continue

    # pyam DataFrame für diese Variable bauen
    long_df = pd.DataFrame({
        'model': [source_model] * len(df_var),
        'scenario': df_var['scenario_fg'] if 'scenario_fg' in df_var.columns else df_var['scenario'],
        'region': df_var['region'],
        'variable': [pyam_var] * len(df_var),
        'unit': [unit] * len(df_var),
        'year': df_var['year'],
        'value': df_var['recycled']  # Passe ggf. die Werte-Spalte an!
    })

    wide_df = long_df.pivot_table(
        index=['model', 'scenario', 'region', 'variable', 'unit'],
        columns='year',
        values='value'
    ).reset_index()

    wide_df.columns = [str(col) if not isinstance(col, int) else str(col) for col in wide_df.columns]
    fixed_cols = ['model', 'scenario', 'region', 'variable', 'unit']
    year_cols = [col for col in wide_df.columns if col not in fixed_cols]
    try:
        year_cols_sorted = sorted(year_cols, key=lambda x: int(x))
    except ValueError:
        year_cols_sorted = year_cols
    final_cols = fixed_cols + year_cols_sorted
    wide_df = wide_df[final_cols]

    all_pyam_dfs.append(wide_df)
    print(f"Konvertiert: {file_name} – {pyam_var} ({len(wide_df)} Zeilen)")

# Export wie gehabt
if all_pyam_dfs:
    result = pd.concat(all_pyam_dfs, ignore_index=True)
    output_path = "output/prism_pyam_export.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    result.to_excel(output_path, index=False)
    print(f"\nErfolgreich exportiert: {output_path} ({len(result)} Zeilen)")
else:
    print("Keine Daten zum Exportieren gefunden.")
