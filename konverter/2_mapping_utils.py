#%%

import pandas as pd
import os
import sys
import time
from collections import defaultdict, Counter
from pathlib import Path

# =========================
# Configuration
# =========================
start_time = time.time()

MAPPING_FILE_PATH = 'variable_mapping_all.xlsx'

# Which column contains the target units in the mapping file
MAPPING_UNIT_COL = "Unit"

# List to collect error/warning messages
error_log = []

# =========================
# Unit-Harmonizer & Collector
# =========================
try:
    import pint
    _UREG = pint.UnitRegistry()
except Exception:
    _UREG = None

# Harmonize unit strings to a common format
_UNIT_SYNONYMS = {
    "tco2": "t CO2",
    "tco2e": "t CO2e",
    "ktco2": "kt CO2",
    "mtco2": "Mt CO2",
    "gwh": "GWh",
    "mwh": "MWh",
    "kwh": "kWh",
    "eur": "EUR",
    "usd": "USD",
    "percent": "%",
    "%": "%",
}

def normalize_unit(u: str) -> str | None:
    if not isinstance(u, str) or not u.strip():
        return None
    s = u.strip()
    key = s.replace(" ", "").lower()
    if key in _UNIT_SYNONYMS:
        s = _UNIT_SYNONYMS[key]
    # "MtCO2" -> "Mt CO2"
    if key.endswith("co2") and "CO2" not in s:
        s = s.replace("co2", "CO2")
        if " " not in s:
            s = s.replace("CO2", " CO2")
    # only validation; CO2 is not a valid unit
    if _UREG:
        try:
            _ = _UREG.parse_expression(s)
        except Exception:
            pass
    return s

class UnitCollector:
    """
    Sammelt Einheiten-Beobachtungen über den gesamten Run.

    Schlüssel: (original_variable, variable_new, model)
    Werte: Counter der beobachteten Units, plus Beispiele & Quellen.
    """
    def __init__(self):
        self._obs = defaultdict(lambda: Counter())
        self._examples = defaultdict(dict)
        self._sources = defaultdict(lambda: defaultdict(int))

    def add(self, original_variable, variable_new, unit, model, *,
            source_file=None, sheet=None, column=None, example_value=None, count: int = 1):
        key = (
            str(original_variable).strip() if original_variable is not None else None,
            str(variable_new).strip() if variable_new is not None else None,
            str(model).strip() if model is not None else None
        )
        unit_norm = normalize_unit(unit) if unit is not None else None
        self._obs[key][unit_norm] += int(count)
        if unit_norm not in self._examples[key] and example_value is not None:
            self._examples[key][unit_norm] = example_value
        if source_file:
            src = f"{Path(source_file).name}:{sheet or ''}:{column or ''}"
            self._sources[key][src] += int(count)

    def to_suggestions_df(self) -> pd.DataFrame:
        rows = []
        for key, cnts in self._obs.items():
            original_variable, variable_new, model = key
            total = sum(cnts.values())
            unit_suggest, freq = (None, 0)
            if cnts:
                unit_suggest, freq = cnts.most_common(1)[0]
            conflict = len([u for u in cnts if u]) > 1
            units_detail = "; ".join([f"{u or 'None'}×{n}" for u, n in cnts.most_common()])
            example = self._examples.get(key, {}).get(unit_suggest)
            sources = ", ".join(sorted(self._sources.get(key, {}).keys()))
            rows.append({
                "Variable value (original)": original_variable,
                "Variable name (new)": variable_new,
                "Source model": model,
                "suggested_unit": unit_suggest,
                "observations": total,
                "units_seen": units_detail,
                "example_value": example,
                "conflict": conflict,
                "sources": sources,
            })
        return pd.DataFrame(rows)

_UNIT_COLLECTOR = UnitCollector()

def record_unit(original_variable, variable_new, unit, model, *,
                source_file=None, sheet=None, column=None, example_value=None, count: int = 1):
    """Einfacher Wrapper, überall im Flow aufrufbar."""
    _UNIT_COLLECTOR.add(
        original_variable=original_variable,
        variable_new=variable_new,
        unit=unit,
        model=model,
        source_file=source_file,
        sheet=sheet,
        column=column,
        example_value=example_value,
        count=count,
    )

def _first_sheet_name(xlsx_path: str) -> str:
    xf = pd.ExcelFile(xlsx_path)
    return xf.sheet_names[0]

def _ensure_columns(df: pd.DataFrame, cols):
    for c in cols:
        if c not in df.columns:
            df[c] = None
    return df

def apply_unit_suggestions(mapping_xlsx: str,
                           mapping_sheet_name: str,
                           unit_col: str = MAPPING_UNIT_COL,
                           observed_sheet: str = "auto_units_observed",
                           conflicts_sheet: str = "auto_units_conflicts",
                           dry_run: bool = False) -> pd.DataFrame:
    """
    Füllt leere `unit_col`-Zellen im Mapping-Sheet automatisch, wenn es
    eine eindeutige Beobachtung (kein Konflikt) gab. Schreibt außerdem
    zwei Hilfs-Sheets für Überblick & Konflikte.
    """
    xls_path = Path(mapping_xlsx)
    try:
        current = pd.read_excel(xls_path, sheet_name=mapping_sheet_name)
    except Exception as e:
        raise RuntimeError(f"Konnte Blatt '{mapping_sheet_name}' in {mapping_xlsx} nicht lesen: {e}")


    key_cols = ["Variable value (original)", "Variable name (new)", "Source model"]
    current = _ensure_columns(current, [*key_cols, unit_col])

    # 🔧 Keys im Mapping-Sheet hart normalisieren
    for col in key_cols:
        current[col] = (
            current[col]
            .astype(str)
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)   # Mehrfachspaces → 1 Space
        )

    # Vorschläge holen
    sugg = _UNIT_COLLECTOR.to_suggestions_df()

    # 🔧 Keys in den Vorschlägen ebenfalls normalisieren
    for col in key_cols:
        sugg[col] = (
            sugg[col]
            .astype(str)
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )

    # (Optional) Case-insensitive match erzwingen:
    # for col in key_cols:
    #     current[col] = current[col].str.casefold()
    #     sugg[col]    = sugg[col].str.casefold()

    # Merge
    merged = current.merge(
        sugg[[*key_cols, "suggested_unit", "conflict", "units_seen", "observations"]],
        on=key_cols, how="left",
    )

    # Nur leere Unit + eindeutige Vorschläge schreiben
    mask_empty = merged[unit_col].isna() | (merged[unit_col].astype(str).str.strip() == "")
    mask_ok = mask_empty & merged["suggested_unit"].notna() & (merged["conflict"] == False)
    num_fill = int(mask_ok.sum())
    merged.loc[mask_ok, unit_col] = merged.loc[mask_ok, "suggested_unit"]

    # 🕵️ Debug: Warum blieb etwas leer?
    if num_fill == 0:
        # Zeilen im Mapping, zu denen es KEINE Vorschläge gab (Left-Anti-Join)
        no_sugg = merged[mask_empty & merged["suggested_unit"].isna()][key_cols].drop_duplicates()
        print("[Hinweis] Keine Übereinstimmung für diese Keys (Top 20):")
        print(no_sugg.head(20).to_string(index=False))


    # Now merge
    merged = current.merge(
        sugg[[*key_cols, "suggested_unit", "conflict", "units_seen", "observations"]],
        on=key_cols,
        how="left",
    )

    mask_empty = merged[unit_col].isna() | (merged[unit_col].astype(str).str.strip() == "")
    mask_ok = mask_empty & merged["suggested_unit"].notna() & (merged["conflict"] == False)
    num_fill = int(mask_ok.sum())
    merged.loc[mask_ok, unit_col] = merged.loc[mask_ok, "suggested_unit"]

    if not dry_run:
        with pd.ExcelWriter(xls_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
            merged[current.columns].to_excel(xw, sheet_name=mapping_sheet_name, index=False)
            sugg.sort_values(["conflict", "Variable name (new)", "Source model"]).to_excel(
                xw, sheet_name=observed_sheet, index=False
            )
            conflicts = sugg[sugg["conflict"] == True].copy()
            if not conflicts.empty:
                conflicts.to_excel(xw, sheet_name=conflicts_sheet, index=False)
        print(f"Mapping aktualisiert: {num_fill} Einheit(en) automatisch gesetzt.")
    else:
        print(f"[Dry-Run] Würde {num_fill} Einheit(en) setzen.")

    return merged[current.columns]


# =========================
# 1. Zentrale Mapping-Dateien einlesen
# =========================
print(f"Lese zentrale Mapping-Datei: {MAPPING_FILE_PATH}")
try:
    # erstes Sheet als DataFrame (enthält die Spalten für Mapping)
    df_mapping_full = pd.read_excel(MAPPING_FILE_PATH, sheet_name='variable_mapping').fillna('')
    # Name des ersten Sheets auch für das spätere Zurückschreiben merken
    FIRST_MAPPING_SHEET_NAME = _first_sheet_name(MAPPING_FILE_PATH)

    # region_mapping (optional)
    try:
        df_region_map = pd.read_excel(MAPPING_FILE_PATH, sheet_name='region_mapping')
        region_mapper = pd.Series(df_region_map.Target_Region.values, index=df_region_map.Source_Region).to_dict()
        print("Regionen-Mapping erfolgreich geladen.")
    except ValueError:
        print("INFO: Kein 'region_mapping'-Tabellenblatt gefunden. Regionen werden nicht umbenannt.")
        error_log.append("INFO: Kein 'region_mapping'-Tabellenblatt gefunden. Regionen werden nicht umbenannt.")
        region_mapper = {}

except FileNotFoundError:
    msg = f"FEHLER: Die Mapping-Datei '{MAPPING_FILE_PATH}' wurde nicht gefunden."
    print(msg)
    error_log.append(msg)
    sys.exit(1)


# =========================
# 2. Gruppiere Mappings nach Zieldatei
# =========================
grouped_mappings = df_mapping_full.groupby(['File location', 'File name', 'Source model'])
print(f"\n{len(grouped_mappings)} einzigartige Dateien zur Verarbeitung gefunden.")


# =========================
# 3. Schleife über jede Datei
# =========================
for (file_location, file_name, model), group_of_mappings in grouped_mappings:
    config = group_of_mappings.iloc[0]
#    INPUT_FILE_PATH = os.path.join('input', file_location, file_name)
    INPUT_FILE_PATH = os.path.join('input\\POC_2.0_2025.10', file_location, file_name)
    output_filename = 'pyam_' + model + '-' + os.path.splitext(file_name)[0] + '.xlsx'
    OUTPUT_FILE_PATH = os.path.join('output', output_filename)

    print(f"\n--- Starte Verarbeitung für: {file_name} ---")
    error_log.append(f"\n--- {file_name} ---")
    # Sheet-Name aus Mapping lesen
    sheet_name = config.get('Sheet name', 0)  # Default: erstes Sheet (Index 0)
    if sheet_name == '' or pd.isna(sheet_name):
        sheet_name = 0

    print(f"Verwende Sheet: {sheet_name}")

    # Datei einlesen mit Sheet-Unterstützung
    try:
        if file_name.lower().endswith('.xlsx'):
            df_input = pd.read_excel(INPUT_FILE_PATH, sheet_name=sheet_name)
        elif file_name.lower().endswith('.csv'):
            separator = config['Separator'] if config['Separator'] else ','
            if '250424' in file_name:
                 df_input = pd.read_csv(INPUT_FILE_PATH, sep=separator, index_col=0, low_memory=False)
            else:
                 df_input = pd.read_csv(INPUT_FILE_PATH, sep=separator, low_memory=False)
        else:
            print(f"WARNUNG: Unbekanntes Format. Übersprungen.")
            error_log.append(f"WARNUNG: Unbekanntes Format. Übersprungen. {file_name}\n" + "-" * 40)
            continue
        print(f"Datei '{INPUT_FILE_PATH}' (Sheet: {sheet_name}) erfolgreich geladen.")
    except Exception as e:
        print(f"FEHLER beim Einlesen: {e}. Übersprungen.")
        error_log.append(f"FEHLER beim Einlesen: {e}. Übersprungen. {file_name}\n" + "-" * 40)
        continue

    # # --- Variable-Mapping ---
    # mapping_source_columns = config['Variable column']
    # try:
    #     df_input.columns = df_input.columns.str.strip()
    #     if '|' in mapping_source_columns:
    #         columns_to_combine = [col.strip() for col in mapping_source_columns.split('|')]
    #         df_input['original_variable'] = df_input[columns_to_combine].astype(str).agg(' | '.join, axis=1)
    #     else:
    #         df_input['original_variable'] = df_input[mapping_source_columns]
    # except KeyError as e:
    #     print(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
    #     error_log.append(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen. {file_name}\n" + "-" * 40)
    #     continue

    # variable_mapper = pd.Series(
    #     group_of_mappings['Variable name (new)'].values,
    #     index=group_of_mappings['Variable value (original)']
    # ).to_dict()
    # df_input['variable'] = df_input['original_variable'].map(variable_mapper)

    # unmapped_mask = df_input['variable'].isna()
    # if unmapped_mask.any():
    #     unique_unmapped_keys = df_input[unmapped_mask]['original_variable'].unique()
    #     print("\nWARNUNG: Folgende Variablen wurden gefunden, aber nicht zugeordnet:")
    #     error_log.append("\nWARNUNG: Folgende Variablen wurden gefunden, aber nicht zugeordnet:")
    #     for key in sorted(list(unique_unmapped_keys)): 
    #         print(f"{key}")
    #         error_log.append(f"{key}")
    #     print("-" * 40)

    # --- Variable-Mapping (robuster) ---
    mapping_source_columns = config['Variable column']

    # Hilfsfunktion: Spalte sicher als String ohne NaN behandeln
    def _to_clean_string(series: pd.Series) -> pd.Series:
        # 'string' Dtype vermeidet gemischte Typen; fillna('') verhindert float('nan')
        return series.astype('string').fillna('').str.strip()

    try:
        # Spaltennamen trimmen
        df_input.columns = df_input.columns.astype('string').str.strip()

        if '|' in mapping_source_columns:
            columns_to_combine = [col.strip() for col in mapping_source_columns.split('|')]

            # Prüfen, ob alle Spalten vorhanden sind
            missing_cols = [c for c in columns_to_combine if c not in df_input.columns]
            if missing_cols:
                raise KeyError(f"{missing_cols}")

            # Jede Spalte zu String + NaN -> '', dann joinen
            cleaned = df_input[columns_to_combine].apply(_to_clean_string)
            df_input['original_variable'] = cleaned.agg(' | '.join, axis=1)
        else:
            if mapping_source_columns not in df_input.columns:
                raise KeyError(mapping_source_columns)
            df_input['original_variable'] = _to_clean_string(df_input[mapping_source_columns])

    except KeyError as e:
        print(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
        error_log.append(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen. {file_name}\n" + "-" * 40)
        continue

    # Mapper bauen – Keys/Values string-normalisieren
    mapper_index = group_of_mappings['Variable value (original)'].astype('string').fillna('').str.strip()
    mapper_values = group_of_mappings['Variable name (new)'].astype('string').fillna('').str.strip()
    variable_mapper = pd.Series(mapper_values.values, index=mapper_index.values).to_dict()

    # original_variable für Mapping und Vergleich ebenfalls trimmen
    df_input['original_variable'] = df_input['original_variable'].astype('string').str.strip()

    # Optional: Case-insensitive (beidseitig auf lower):
    # df_input['original_variable'] = df_input['original_variable'].str.lower()
    # variable_mapper = {str(k).lower(): v for k, v in variable_mapper.items()}

    df_input['variable'] = df_input['original_variable'].map(variable_mapper)

    # Unmapped-Keys robust ermitteln und sortieren (nur Strings, keine NaN/Floats)
    unmapped_mask = df_input['variable'].isna()
    if unmapped_mask.any():
        unique_unmapped_keys = (
            df_input.loc[unmapped_mask, 'original_variable']
            .dropna()
            .astype('string')
            .str.strip()
            .unique()
            .tolist()
        )
        unique_unmapped_keys = [str(k) for k in unique_unmapped_keys if str(k) != '']
        print("\nWARNUNG: Folgende Variablen wurden gefunden, aber nicht zugeordnet:")
        error_log.append("\nWARNUNG: Folgende Variablen wurden gefunden, aber nicht zugeordnet:")
        for key in sorted(unique_unmapped_keys):
            print(f"{key}")
            error_log.append(f"{key}")
        print("-" * 40)

    
    df_input.dropna(subset=['variable'], inplace=True)
    if df_input.empty:
        print("INFO: Keine gültigen Daten nach Mapping. Keine Ausgabe erstellt.")
        error_log.append(f"INFO: Keine gültigen Daten nach Mapping. Keine Ausgabe erstellt. {file_name}")
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

        # Source model aus Mapping lesen
        if 'Source model' in config and config['Source model']:
            if config['Source model'] in df_input.columns:
                df_iamc['model'] = df_input[config['Source model']]
            else:
                df_iamc['model'] = config['Source model']  # fester Wert
        else:
            df_iamc['model'] = 'Unknown Model'

        # Unit-Handling (aus Quelle oder fixer Wert)
        if config['Source Unit']:
            df_iamc['unit'] = df_input[config['Source Unit']]
        elif 'Unit' in config and config['Unit']:
            df_iamc['unit'] = config['Unit']
        else:
            df_iamc['unit'] = 'undefined'

        # --- Regionen-Mapping ---
        if 'region' in df_iamc.columns and isinstance(region_mapper, dict) and region_mapper:
            found_regions = set(df_iamc['region'].unique())
            mappable_regions = set(region_mapper.keys())
            unmapped_regions = found_regions.difference(mappable_regions)

            if unmapped_regions:
                print("\nWARNUNG: Folgende Regionen wurden gefunden, aber nicht im 'region_mapping' definiert:")
                print("Diese Regionen werden im Original beibehalten. Fügen Sie bei Bedarf Mappings hinzu:")
                error_log.append("\nWARNUNG: Folgende Regionen wurden gefunden, aber nicht im 'region_mapping' definiert:")
                error_log.append("Diese Regionen werden im Original beibehalten. Fügen Sie bei Bedarf Mappings hinzu:")
                if unmapped_regions == "nan":
                    print("Es sind Zeilen ohne Angabe einer Region enthalten")
                for region_code in sorted(list(unmapped_regions)):
                    print(f"{region_code}")
                    error_log.append(f"{region_code}")
                print("-" * 40)

            # Mapping anwenden
            df_iamc['region'] = df_iamc['region'].map(region_mapper).fillna(df_iamc['region'])
            print("Regionen-Mapping wurde angewendet.")

    except KeyError as e:
        print(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
        error_log.append(f"FEHLER: Spalte {e} nicht gefunden. Übersprungen.")
        continue

    # --- NEU: Einheiten sammeln (effizient, ohne jede Zeile einzeln zu iterieren) ---
    # Wir verdichten auf Kombinationen und zählen deren Auftreten
    try:
        # Beispielwert (erste vorhandene value-Zelle) – nur für Transparenz im Report
        example_val = None
        try:
            example_val = df_iamc['value'].dropna().iloc[0]
        except Exception:
            pass

        obs_df = pd.DataFrame({
            "original_variable": df_input['original_variable'].values,
            "variable_new": df_iamc['variable'].values,
            "model": df_iamc['model'].values,
            "unit": df_iamc['unit'].values,
        })

        grp = obs_df.groupby(["original_variable", "variable_new", "model", "unit"]).size().reset_index(name="n")
        for _, row in grp.iterrows():
            record_unit(
                original_variable=row["original_variable"],
                variable_new=row["variable_new"],
                unit=row["unit"],
                model=row["model"],
                source_file=INPUT_FILE_PATH,
                sheet=sheet_name,
                column=config.get('Variable column'),
                example_value=example_val,
                count=int(row["n"]),
            )
    except Exception as e:
        print(f"WARNUNG: Konnte Einheiten-Beobachtungen nicht sammeln: {e}")

    # --- Pivotieren & Speichern ---
    print("Daten werden pivotiert...")
    try:
        df_output = df_iamc.pivot_table(
            index=['model', 'scenario', 'region', 'variable', 'unit'],
            columns='year',
            values='value',
            aggfunc='sum'
        ).reset_index()
        df_output.columns = [str(col) for col in df_output.columns]
    except Exception as e:
        print(f"FEHLER während des Pivotierens: {e}")
        error_log.append(f"FEHLER während des Pivotierens: {e}")
        continue

    os.makedirs(os.path.dirname(OUTPUT_FILE_PATH), exist_ok=True)
    df_output.to_excel(OUTPUT_FILE_PATH, index=False, sheet_name='pyam_data')
    print(f"Verarbeitung abgeschlossen. Ergebnis in: {OUTPUT_FILE_PATH}")


# =========================
# Nachlauf: Mapping-Excel automatisch mit Units anreichern
# =========================
try:
    apply_unit_suggestions(
        mapping_xlsx=MAPPING_FILE_PATH,
        mapping_sheet_name="variable_mapping",  # erstes Sheet überschreiben
        unit_col=MAPPING_UNIT_COL,
        observed_sheet="auto_units_observed",
        conflicts_sheet="auto_units_conflicts",
        dry_run=False,  # zuerst gern auf True stellen, um zu sehen, was passieren würde
    )
except RuntimeError as e:
    print(f"Hinweis: {e}\n"
          f"(Stelle sicher, dass das erste Sheet Spalten "
          f"'Variable value (original)', 'Variable name (new)', "
          f"'Source model' und '{MAPPING_UNIT_COL}' enthält.)")


# =========================
# Abschluss
# =========================
print("\n\nAlle Dateien aus der Mapping-Tabelle wurden verarbeitet.")
print(error_log)

# Exportiere das error_log als Textdatei
with open("output/error_log.txt", "w", encoding="utf-8") as f:
    for line in error_log:
        f.write(str(line) + "\n")

end_time = time.time()
elapsed = end_time - start_time
print(f"\n⏱️ Laufzeit des Skripts: {elapsed:.2f} Sekunden")
