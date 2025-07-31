#%%

import pandas as pd
import os
import sys
import time

start_time = time.time()

# --- Konfiguration ---
# MAPPING_FILE_PATH = 'variable_mapping_CITS_PRISM_TIMES.xslx'
# MAPPING_FILE_PATH = 'variable_mapping_TIAM.xlsx'
# MAPPING_FILE_PATH = 'variable_mapping_exiomod_globiom.xlsx'
# MAPPING_FILE_PATH = 'variable_mapping_TIMEs2.xlsx'
MAPPING_FILE_PATH = 'variable_mapping_all.xlsx'

# Liste zur Zwischenspeicherung aller Fehlermeldungen
error_log = []  


# --- 1. Zentrale Mapping-Dateien einlesen ---
print(f"Lese zentrale Mapping-Datei: {MAPPING_FILE_PATH}")
try:
    df_mapping_full = pd.read_excel(MAPPING_FILE_PATH, sheet_name=0).fillna('') 
    df_region_map = pd.read_excel(MAPPING_FILE_PATH, sheet_name='region_mapping')
    region_mapper = pd.Series(df_region_map.Target_Region.values, index=df_region_map.Source_Region).to_dict()
    print("Regionen-Mapping erfolgreich geladen.")
except FileNotFoundError:
    print(f"FEHLER: Die Mapping-Datei '{MAPPING_FILE_PATH}' wurde nicht gefunden.")
    error_log.append(f"FEHLER: Die Mapping-Datei '{MAPPING_FILE_PATH}' wurde nicht gefunden.")

    sys.exit(1)
except ValueError:
    print("INFO: Kein 'region_mapping'-Tabellenblatt gefunden. Regionen werden nicht umbenannt.")
    error_log.append(f"INFO: Kein 'region_mapping'-Tabellenblatt gefunden. Regionen werden nicht umbenannt.")
    region_mapper = {}

# --- 2. Gruppiere Mappings nach Zieldatei ---
grouped_mappings = df_mapping_full.groupby(['File location', 'File name', 'Source model'])
print(f"\n{len(grouped_mappings)} einzigartige Dateien zur Verarbeitung gefunden.")
# --- 3. Schleife über jede Datei ---
for (file_location, file_name, model), group_of_mappings in grouped_mappings:
    # print (group_of_mappings)
    config = group_of_mappings.iloc[0]
    INPUT_FILE_PATH = os.path.join('input', file_location, file_name)
    output_filename = 'pyam_' + model + '-' + os.path.splitext(file_name)[0] + '.xlsx'
    OUTPUT_FILE_PATH = os.path.join('output', output_filename)

    print(f"\n--- Starte Verarbeitung für: {file_name} ---")

    # --- NEU: Sheet-Name aus Mapping lesen ---
    sheet_name = config.get('Sheet name', 0)  # Default: erstes Sheet (Index 0)
    if sheet_name == '' or pd.isna(sheet_name):
        sheet_name = 0  # Falls leer, verwende erstes Sheet
    
    print(f"Verwende Sheet: {sheet_name}")

    # --- Datei einlesen mit Sheet-Unterstützung ---
    try:
        if file_name.lower().endswith('.xlsx'):
            df_input = pd.read_excel(INPUT_FILE_PATH, sheet_name=sheet_name)
        elif file_name.lower().endswith('.csv'):
            separator = config['Separator'] if config['Separator'] else ','
            if '250424' in file_name:
                 df_input = pd.read_csv(INPUT_FILE_PATH, sep=separator, index_col=0)
            else:
                 df_input = pd.read_csv(INPUT_FILE_PATH, sep=separator)
        else:
            print(f"WARNUNG: Unbekanntes Format. Übersprungen.")
            error_log.append(f"WARNUNG: Unbekanntes Format. Übersprungen." + file_name +"\n" + "-" * 40)
            continue
        print(f"Datei '{INPUT_FILE_PATH}' (Sheet: {sheet_name}) erfolgreich geladen.")
    except Exception as e:
        print(f"FEHLER beim Einlesen: {e}. Übersprungen.")
        error_log.append("FEHLER beim Einlesen: {e}. Übersprungen." + file_name +"\n" + "-" * 40)
        continue

    # --- Variable-Mapping (bleibt gleich) ---
    mapping_source_columns = config['Variable column']
    try:
        df_input.columns = df_input.columns.str.strip()
        if '|' in mapping_source_columns:
            columns_to_combine = [col.strip() for col in mapping_source_columns.split('|')]
            df_input['original_variable'] = df_input[columns_to_combine].astype(str).agg(' | '.join, axis=1)
        else:
            df_input['original_variable'] = df_input[mapping_source_columns]
    except KeyError as e:
        print(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
        error_log.append(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen." + file_name +"\n" + "-" * 40)
        continue

    variable_mapper = pd.Series(group_of_mappings['Variable name (new)'].values, index=group_of_mappings['Variable value (original)']).to_dict()
    df_input['variable'] = df_input['original_variable'].map(variable_mapper)
    
    unmapped_mask = df_input['variable'].isna()
    if unmapped_mask.any():
        unique_unmapped_keys = df_input[unmapped_mask]['original_variable'].unique()
        print("\nWARNUNG: Folgende Variablen wurden gefunden, aber nicht zugeordnet:")
        error_log.append("\nWARNUNG: Folgende Variablen wurden gefunden, aber nicht zugeordnet:")
        for key in sorted(list(unique_unmapped_keys)): print(f"{key}")
        for key in sorted(list(unique_unmapped_keys)): error_log.append(f"{key}")
        print("-" * 40)

    df_input.dropna(subset=['variable'], inplace=True)
    if df_input.empty:
        print("INFO: Keine gültigen Daten nach Mapping. Keine Ausgabe erstellt.")
        error_log.append("INFO: Keine gültigen Daten nach Mapping. Keine Ausgabe erstellt.",file_name)
        continue

    # --- Transformation ---
    print("Daten werden für das Pivotieren vorbereitet...")
    try:
        data_for_iamc = {
            'scenario': df_input[config['Source Scenario']],
            'region': df_input[config['Source Region']],
            'year': df_input[config['Source Year']],
            'value': df_input[config['Source Value']],
            'variable': df_input['variable']
        }
        df_iamc = pd.DataFrame(data_for_iamc)
        
        # --- NEU: Source model aus Mapping lesen ---
        if 'Source model' in config and config['Source model']:
            if config['Source model'] in df_input.columns:
                df_iamc['model'] = df_input[config['Source model']]
            else:
                df_iamc['model'] = config['Source model']  # Als fester Wert verwenden
        else:
            df_iamc['model'] = 'Unknown Model'  # Default-Wert
        
        # Unit-Handling (bleibt gleich)
        if config['Source Unit']:
            df_iamc['unit'] = df_input[config['Source Unit']]
        elif 'Unit' in config and config['Unit']:
            df_iamc['unit'] = config['Unit']
        else:
            df_iamc['unit'] = 'undefined'
            
        # --- Regionen-Mapping (bleibt gleich) ---
        if region_mapper:
            found_regions = set(df_iamc['region'].unique())
            mappable_regions = set(region_mapper.keys())
            unmapped_regions = found_regions.difference(mappable_regions)

            if unmapped_regions:
                print("\nWARNUNG: Folgende Regionen wurden gefunden, aber nicht im 'region_mapping' definiert:")
                print("Diese Regionen werden im Original beibehalten. Fügen Sie bei Bedarf Mappings hinzu:")
                error_log.append("\nWARNUNG: Folgende Regionen wurden gefunden, aber nicht im 'region_mapping' definiert:")
                error_log.append("Diese Regionen werden im Original beibehalten. Fügen Sie bei Bedarf Mappings hinzu:")
                for region_code in sorted(list(unmapped_regions)):
                    print(f"{region_code}")
                    error_log.append(f"{region_code}")
                print("-" * 40)
        
        # Wende das Regionen-Mapping an
        if region_mapper:
            df_iamc['region'] = df_iamc['region'].map(region_mapper).fillna(df_iamc['region'])
            print("Regionen-Mapping wurde angewendet.")

    except KeyError as e:
         print(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
         error_log.append(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
         continue
    
    # --- Pivotieren & Speichern (bleibt gleich) ---
    print("Daten werden pivotiert...")
    try:
        df_output = df_iamc.pivot_table(index=['model', 'scenario', 'region', 'variable', 'unit'], columns='year', values='value', aggfunc='sum').reset_index()
        df_output.columns = [str(col) for col in df_output.columns]
    except Exception as e:
        print(f"FEHLER während des Pivotierens: {e}")
        error_log.append(f"FEHLER während des Pivotierens: {e}")
        continue

    os.makedirs(os.path.dirname(OUTPUT_FILE_PATH), exist_ok=True)
    df_output.to_excel(OUTPUT_FILE_PATH, index=False, sheet_name='pyam_data')
    print(f"Verarbeitung abgeschlossen. Ergebnis in: {OUTPUT_FILE_PATH}")
    # print("\nVorschau der ersten 5 Zeilen der erstellten Datei:")
    # print(df_output.head().to_string())

print("\n\nAlle Dateien aus der Mapping-Tabelle wurden verarbeitet.")
print(error_log)

end_time = time.time()
elapsed = end_time - start_time
print(f"\n⏱️ Laufzeit des Skripts: {elapsed:.2f} Sekunden")