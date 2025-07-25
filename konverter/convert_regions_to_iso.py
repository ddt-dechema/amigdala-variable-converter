import pandas as pd
import country_converter as coco

def convert_regions_to_csv(file_path, sheet_name='region_mapping'):
    """
    Liest Excel-Datei und erstellt eine CSV mit Source_Region, ISO2 und ISO3 Spalten
    
    Parameters:
    file_path: Pfad zur Excel-Datei
    sheet_name: Name des Arbeitsblatts
    """
    
    # Lade die Excel-Datei
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Excel-Datei erfolgreich geladen. {len(df)} Zeilen gefunden.")
    except Exception as e:
        print(f"Fehler beim Laden der Datei: {e}")
        return None
    
    # Überprüfe ob die erforderlichen Spalten vorhanden sind
    if 'Source_Region' not in df.columns:
        print("Fehler: Spalte 'Source_Region' nicht gefunden!")
        return None
    
    # Initialisiere den Country Converter
    cc = coco.CountryConverter()
    
    # Erstelle neue Spalten mit ISO2 und ISO3 Codes
    print("Konvertiere zu ISO2-Codes...")
    df['ISO2'] = cc.convert(
        names=df['Source_Region'].fillna(''), 
        to='ISO2',
        not_found=None
    )
    
    print("Konvertiere zu ISO3-Codes...")
    df['ISO3'] = cc.convert(
        names=df['Source_Region'].fillna(''), 
        to='ISO3',
        not_found=None
    )
    
    # Erstelle finale Tabelle mit nur den gewünschten Spalten
    result_df = df[['Source_Region', 'ISO2', 'ISO3']].copy()
    
    # Zeige Ergebnisse
    print(f"\nKonvertierung abgeschlossen!")
    print("\nErste 10 Zeilen:")
    print(result_df.head(10))
    
    # Zeige nicht konvertierte Einträge
    not_converted = result_df[
        ((result_df['ISO2'].isna()) | (result_df['ISO3'].isna())) & 
        (result_df['Source_Region'].notna()) & 
        (result_df['Source_Region'] != '')
    ]
    
    if not not_converted.empty:
        print(f"\nWarnung: Einige Einträge konnten nicht konvertiert werden:")
        for country in not_converted['Source_Region'].unique():
            print(f"  - '{country}'")
    
    # Zeige Statistiken
    total_entries = len(result_df[result_df['Source_Region'].notna() & (result_df['Source_Region'] != '')])
    converted_iso2 = len(result_df[result_df['ISO2'].notna()])
    converted_iso3 = len(result_df[result_df['ISO3'].notna()])
    
    print(f"\nStatistiken:")
    print(f"  - Gesamte Einträge: {total_entries}")
    print(f"  - ISO2 erfolgreich: {converted_iso2} ({converted_iso2/total_entries*100:.1f}%)")
    print(f"  - ISO3 erfolgreich: {converted_iso3} ({converted_iso3/total_entries*100:.1f}%)")
    
    # Speichere als CSV
    output_file = file_path.replace('.xlsx', '_iso_codes.csv')
    result_df.to_csv(output_file, index=False, encoding='utf-8')
    print(f"\nErgebnisse gespeichert als CSV: {output_file}")
    
    return result_df

def test_conversion_examples():
    """Testet die Konvertierung mit einigen Beispielen"""
    cc = coco.CountryConverter()
    
    test_countries = [
        'Germany', 'USA', 'United Kingdom', 'China', 'Russia', 
        'South Korea', 'Iran', 'Venezuela', 'Czech Republic'
    ]
    
    print("Test-Konvertierungen:")
    print("Land\t\t\tISO2\tISO3")
    print("-" * 50)
    
    for country in test_countries:
        iso2 = cc.convert(country, to='ISO2')
        iso3 = cc.convert(country, to='ISO3')
        print(f"{country:<20}\t{iso2}\t{iso3}")

# Hauptskript
if __name__ == "__main__":
    print("=== Länder zu ISO-Codes Konverter ===\n")
    
    # Teste einige Konvertierungen als Beispiel
    test_conversion_examples()
    
    print("\n" + "="*60)
    
    # Pfad zu deiner Excel-Datei
    file_path = "variable_mapping_new.xlsx"
    
    # Konvertiere und speichere als CSV
    print(f"\nVerarbeite Datei: {file_path}")
    df_result = convert_regions_to_csv(file_path)
    
    if df_result is not None:
        print("\n✅ Fertig! Die CSV-Datei enthält:")
        print("   - Source_Region: Original-Ländernamen")
        print("   - ISO2: 2-stellige ISO-Codes (DE, US, etc.)")
        print("   - ISO3: 3-stellige ISO-Codes (DEU, USA, etc.)")
