import pandas as pd
import os
import sys

# --- Konfiguration ---
# Es gibt keine Zieldatei mehr, das Skript verarbeitet alle Dateien aus der Mapping-Datei.
MAPPING_FILE_PATH = 'variable_mapping.xlsx'

# --- 1. Zentrale Mapping-Datei einlesen ---
print(f"Lese zentrale Mapping-Datei: {MAPPING_FILE_PATH}")
try:
    df_mapping_full = pd.read_excel(MAPPING_FILE_PATH)
except FileNotFoundError:
    print(f"FEHLER: Die Mapping-Datei '{MAPPING_FILE_PATH}' wurde nicht gefunden. Skript wird beendet.")
    sys.exit(1)

# --- 2. NEU: Gruppiere alle Mapping-Regeln nach der Datei, für die sie gelten ---
# Das stellt sicher, dass jede Input-Datei nur einmal geöffnet wird.
grouped_mappings = df_mapping_full.groupby(['File location', 'File name'])

print(f"\n{len(grouped_mappings)} einzigartige Dateien zur Verarbeitung in der Mapping-Datei gefunden.")

# --- 3. NEU: Schleife über jede zu verarbeitende Datei ---
for (file_location, file_name), group_of_mappings in grouped_mappings:
    
    # Konstruiere die Pfade für die aktuelle Datei
    INPUT_FILE_PATH = os.path.join('input', file_location, file_name)
    output_filename = file_name.replace('.xlsx', '_pyam.xlsx').replace('.csv', '_pyam.xlsx')
    OUTPUT_FILE_PATH = os.path.join('output', output_filename)
    
    # Extrahiere Modellname und Mapping-Anweisung aus der ersten Zeile der Gruppe
    model_name = group_of_mappings['Source model'].iloc[0]
    mapping_source_columns = group_of_mappings['Variable column'].iloc[0]

    print(f"\n--- Starte Verarbeitung für: {file_name} ---")
    print(f"Modell: '{model_name}', Mapping-Basis: '{mapping_source_columns}'")

    # --- Daten für die aktuelle Datei einlesen ---
    try:
        df_input = pd.read_excel(INPUT_FILE_PATH)
        print(f"Datei '{INPUT_FILE_PATH}' erfolgreich geladen.")
    except FileNotFoundError:
        print(f"FEHLER: Die Input-Datei '{INPUT_FILE_PATH}' wurde nicht gefunden. Diese Datei wird übersprungen.")
        continue # Macht mit der nächsten Datei in der Schleife weiter

    # --- Dynamischen Suchschlüssel erstellen ---
    if '|' in mapping_source_columns:
        columns_to_combine = [col.strip() for col in mapping_source_columns.split('|')]
        df_input['original_variable'] = df_input[columns_to_combine].astype(str).agg(' | '.join, axis=1)
    else:
        df_input['original_variable'] = df_input[mapping_source_columns]

    # --- Mapping anwenden ---
    variable_mapper = pd.Series(
        group_of_mappings['Variable name (new)'].values,
        index=group_of_mappings['Variable value (original)']
    ).to_dict()
    
    df_input['variable'] = df_input['original_variable'].map(variable_mapper)

    # --- NEU: Finde und melde nicht zugeordnete Variablen ---
    unmapped_mask = df_input['variable'].isna()
    if unmapped_mask.any():
        unique_unmapped_keys = df_input[unmapped_mask]['original_variable'].unique()
        print("\nWARNUNG: Die folgenden Variablen-Kombinationen wurden in der Datei gefunden,")
        print("aber es existiert kein passender Eintrag in 'variable_mapping.xlsx'.")
        print("Diese Zeilen werden ignoriert. Bitte fügen Sie bei Bedarf Mappings hinzu:")
        for key in unique_unmapped_keys:
            print(f"  - {key}")
        print("-" * 20)

    # Entferne die nicht gemappten Zeilen
    df_input.dropna(subset=['variable'], inplace=True)

    if df_input.empty:
        print("INFO: Nach dem Filtern sind keine Daten mehr übrig. Für diese Datei wird keine Ausgabe erstellt.")
        continue

    # --- IAMC-Standardformat vorbereiten und pivotieren ---
    print("Daten werden transformiert und pivotiert...")
    df_iamc = df_input[['Scenario', 'Region', 'variable', 'Unit', 'Period', 'value']].copy()
    df_iamc.rename(columns={'Scenario': 'scenario', 'Region': 'region', 'Unit': 'unit', 'Period': 'year'}, inplace=True)
    df_iamc['model'] = model_name
    
    df_output = df_iamc.pivot_table(
        index=['model', 'scenario', 'region', 'variable', 'unit'],
        columns='year',
        values='value'
    ).reset_index()
    df_output.columns = [int(col) if isinstance(col, float) else col for col in df_output.columns]

    # --- Ergebnis als Excel-Datei speichern ---
    os.makedirs(os.path.dirname(OUTPUT_FILE_PATH), exist_ok=True)
    df_output.to_excel(OUTPUT_FILE_PATH, index=False, sheet_name='pyam_data')
    print(f"Verarbeitung abgeschlossen. Ergebnis gespeichert in: {OUTPUT_FILE_PATH}")

print("\n\nAlle Dateien aus der Mapping-Tabelle wurden verarbeitet.")
