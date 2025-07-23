import pandas as pd
import os

# 1. Mapping laden
mapping_path = "variable_mapping.xlsx"
mapping = pd.read_excel(mapping_path)

model_filter = "PRISM"  # Passe das Modell hier an

filtered_mapping = mapping[mapping['Source model'] == model_filter]

# 2. Basis-Ordner für Input-Dateien
input_base = "input/"

# 3. Liste für alle pyam-DataFrames
all_pyam_dfs = []
for idx, row in mapping.iterrows():
    file_location = row['File location'] if pd.notna(row['File location']) else ''
    file_name = row['File name']
    source_model = row['Source model']
    orig_var = row['Variable name (original)']
    pyam_var = row['Variable name (new)'] if pd.notna(row['Variable name (new)']) and row['Variable name (new)'] else orig_var
    unit = row['Unit'] if pd.notna(row['Unit']) else ''
    definition = row['Definition'] if 'Definition' in row and pd.notna(row['Definition']) else ''
    used_by = row['used by'] if 'used by' in row and pd.notna(row['used by']) else ''

    # 4. Vollständigen Pfad zur Datei bauen
    full_path = os.path.join(input_base, file_location, file_name)

    if not os.path.isfile(full_path):
        print(f"Datei nicht gefunden: {full_path}")
        continue

    # 5. Datei laden (CSV oder Excel)
    if full_path.lower().endswith('.csv'):
        df = pd.read_csv(full_path)
    elif full_path.lower().endswith('.xlsx') or full_path.lower().endswith('.xls'):
        df = pd.read_excel(full_path)
    else:
        print(f"Nicht unterstütztes Dateiformat: {full_path}")
        continue

    # 6. Die Spalte mit den Werten finden (hier: Annahme "recycled" oder orig_var)
    value_col = orig_var if orig_var in df.columns else 'recycled'  # ggf. anpassen je nach Datei
    if value_col not in df.columns:
        print(f"Wert-Spalte '{value_col}' nicht gefunden in {file_name}")
        continue

    # DataFrame ins long-Format
    long_df = pd.DataFrame({
        'model': [source_model] * len(df),
        'scenario': df['scenario_fg'] if 'scenario_fg' in df.columns else df['scenario'],
        'region': df['region'],
        'variable': [pyam_var] * len(df),
        'unit': [unit] * len(df),
        'year': df['year'],
        'value': df[value_col]
    })

    # Wide-Format (Pivot)
    wide_df = long_df.pivot_table(
        index=['model', 'scenario', 'region', 'variable', 'unit'],
        columns='year',
        values='value'
    ).reset_index()

    # Spaltennamen als Strings
    wide_df.columns = [str(col) if not isinstance(col, int) else str(col) for col in wide_df.columns]

    # Liste der gewünschten festen Spalten
    fixed_cols = ['model', 'scenario', 'region', 'variable', 'unit']

    # Finde alle Jahrgangsspalten (alles, was nicht in fixed_cols ist)
    year_cols = [col for col in wide_df.columns if col not in fixed_cols]

    # Sortiere die Jahrgangsspalten numerisch (falls nötig)
    try:
        year_cols_sorted = sorted(year_cols, key=lambda x: int(x))
    except ValueError:
        year_cols_sorted = year_cols  # falls keine reine Jahreszahl, keine Sortierung

    # Setze die finale Spaltenreihenfolge
    final_cols = fixed_cols + year_cols_sorted

    # wende die Reihenfolge an
    wide_df = wide_df[final_cols]

    all_pyam_dfs.append(wide_df)
    print(f"Konvertiert: {file_name} ({len(wide_df)} Zeilen)")

# 8. Alle pyam-DataFrames zusammenführen und speichern
if all_pyam_dfs:
    result = pd.concat(all_pyam_dfs, ignore_index=True)
    output_path = "output/prism_pyam_export.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    result.to_excel(output_path, index=False)
    print(f"\nErfolgreich exportiert: {output_path} ({len(result)} Zeilen)")
else:
    print("Keine Daten zum Exportieren gefunden.")