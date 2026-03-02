import pandas as pd
import country_converter as coco
from pathlib import Path
from config import DICTIONARY_FILE_PATH
# 1. Custom-Mapping
custom_regions = {
    "ACE": "Asia (Eastern)",
    "AEA": "Africa (Eastern)",
    "ASE": "Asia (Southeast)",
    "ASO": "Asia (Southern)",
    "AWE": "Africa (Western)",
    "CHI": "China",
    "EUR": "Europe",
    "NAM": "North America",
    "LAM": "Latin America",
    "MEA": "Middle East & Africa"
}

# 2. converter-Objekt + Funktion
cc = coco.CountryConverter()

def map_region_name(source_name):
    """
    Prüft zuerst eigene Mapping-Tabelle (custom_regions),
    dann country_converter, sonst None.
    """
    if not isinstance(source_name, str) or not source_name.strip():
        return None

    name = source_name.strip()

    # 1. Eigene Mapping-Tabelle prüfen
    if name in custom_regions:
        return custom_regions[name]

    # 2. Mit country_converter versuchen
    result = cc.convert(name, to='name_short', not_found=None)
    if result and isinstance(result, str) and result.lower() != 'not found':
        return result

    # 3. Kein Treffer
    return None

def convert_regions_to_fullname(file_path, sheet_name='regions'):
    """
    Liest Excel-Datei und konvertiert alle Regionen in der Spalte 'source_region'
    zu ausgeschriebenen Ländernamen.
    Speichert das Ergebnis als neue Datei: <originalname>_regions_fullname.xlsx
    """

    # --- Load file ---
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"✅ Datei '{file_path}' erfolgreich geladen ({len(df)} Zeilen).")
    except Exception as e:
        print(f"❌ FEHLER beim Laden der Datei: {e}")
        return

    if 'source_region' not in df.columns:
        print("❌ FEHLER: Spalte 'source_region' nicht gefunden!")
        return


    print("🌍 Konvertiere 'source_region' → ausgeschriebene Ländernamen ('full name') ...")
    
    df['target_region'] = df['source_region'].apply(map_region_name)

    # --- Reporting ---
    total = len(df[df['source_region'].notna() & (df['source_region'] != '')])
    converted = len(df[df['target_region'].notna()])

    print(f"\n📊 Statistik:")
    print(f"  • Gesamt: {total}")
    print(f"  • Erfolgreich konvertiert: {converted} ({converted/total*100 if total>0 else 0:.1f}%)")

    missing = df[df['target_region'].isna() & df['source_region'].notna()]
    if not missing.empty:
        print(f"\n⚠️  WARNUNG: {len(missing)} Regionen konnten nicht zugeordnet werden:")
        for name in sorted(missing['source_region'].unique()):
            print(f"  - {name}")

    # --- Save output as NEW file ---
    input_path = Path(file_path)
    output_file = input_path.with_name(input_path.stem + "_regions_fullname.xlsx")

    df.to_excel(output_file, sheet_name='regions_fullname', index=False)

    print(f"\n💾 Neue Datei gespeichert als: {output_file}\n")
    print("✅ Fertig!")

if __name__ == "__main__":
    print("=== Regionen-Umbenennung zu ausgeschriebenen Ländernamen ===\n")
    # file_path = "dictionary_dataexplorer_variables_translation.xlsm"  # oder deine gewünschte Datei
    file_path = DICTIONARY_FILE_PATH  # aus config.py importieren
    convert_regions_to_fullname(file_path)

