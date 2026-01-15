import os
import re
import pandas as pd
from pathlib import Path

# -------- Konfigurierbare Parameter --------
# Basispfad: Ordner des Skripts
BASE_DIR = Path(__file__).resolve().parent

# Eingabe/ Ausgabe relativ zum Skript (eine Ebene über BASE_DIR)
INPUT_DIR = BASE_DIR.parent / 'input' / 'POC_2.0_2025.10'
OUTPUT_EXCEL = BASE_DIR.parent / 'overview_input_files_unsorted.xlsx'
# ------------------------------------------

def get_excel_sheets_and_columns(filepath: Path):
    """Liest alle Sheetnamen und Spaltenüberschriften aus einer Excel-Datei."""
    try:
        xls = pd.ExcelFile(filepath)
        sheet_info = []
        for sheet in xls.sheet_names:
            try:
                df = xls.parse(sheet, nrows=1)
                cols = ', '.join([str(col) for col in df.columns])
            except Exception:
                cols = ''
            sheet_info.append((sheet, cols))
        return sheet_info
    except Exception:
        return []

def get_csv_columns(filepath: Path):
    """
    Liest die erste Zeile einer CSV-Datei, um Spaltennamen zu ermitteln.
    Robuste Fallbacks für Delimiter und Encoding.
    """
    seps = [None, ';', ',']
    encodings = [None, 'utf-8-sig', 'latin-1']

    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(filepath, nrows=1, sep=sep, engine='python', encoding=enc)
                return ', '.join([str(col) for col in df.columns])
            except Exception:
                continue
    return ''

def derive_source_model(input_dir: Path, file_root: Path) -> str:
    """
    Bestimmt 'Source model' als den obersten Unterordner direkt unter INPUT_DIR.
    Beispiel:
    INPUT_DIR / '1b TIMES' / 'data' / 'file.csv' -> Source model = '1b TIMES'
    Liegt die Datei direkt im INPUT_DIR, wird der Ordnername von INPUT_DIR genutzt.
    """
    try:
        rel = file_root.relative_to(input_dir)
    except Exception:
        # Falls relative_to fehlschlägt, fallback auf Ordnername
        return file_root.name

    parts = rel.parts
    if len(parts) >= 1:
        return parts[0]
    else:
        return input_dir.name

def main():
    print(f'INPUT_DIR: {INPUT_DIR}')
    if not INPUT_DIR.is_dir():
        print(f'✗ Eingabeverzeichnis nicht gefunden: {INPUT_DIR}')
        parent = INPUT_DIR.parent
        if parent.is_dir():
            print('Inhalt von Oberordner:')
            for p in sorted(parent.iterdir()):
                print('-', p.name, '(DIR)' if p.is_dir() else '')
        return

    rows = []

    for root, dirs, files in os.walk(INPUT_DIR):
        root_path = Path(root)
        # Relativer Pfad vom INPUT_DIR
        try:
            rel_root = root_path.relative_to(INPUT_DIR)
            file_location = str(rel_root) if str(rel_root) != '.' else ''
        except Exception:
            file_location = ''

        source_model = derive_source_model(INPUT_DIR, root_path)

        for file in files:
            path = root_path / file
            ext = path.suffix.lower()

            # Nur CSV und Excel
            if ext not in ['.csv', '.xls', '.xlsx']:
                continue

            try:
                if ext in ['.xls', '.xlsx']:
                    sheets = get_excel_sheets_and_columns(path)
                    for sheet, columns in sheets:
                        rows.append({
                            'Source model': source_model,
                            'File location': file_location,
                            'File name': path.name,
                            'Sheet name': sheet,
                            'Column names': columns
                        })
                elif ext == '.csv':
                    columns = get_csv_columns(path)
                    rows.append({
                        'Source model': source_model,
                        'File location': file_location,
                        'File name': path.name,
                        'Sheet name': '',
                        'Column names': columns
                    })
            except Exception:
                # Einzeldateiprobleme ignorieren
                continue

    if not rows:
        print('⚠️ Keine passenden Dateien gefunden oder keine lesbaren Spalten ermittelt. Es wurde keine Excel-Datei geschrieben.')
        return

    # Export
    try:
        df = pd.DataFrame(rows)
        # Spalten explizit ordnen
        cols = ['Source model', 'File location', 'File name', 'Sheet name', 'Column names']
        df = df[cols]
        df.to_excel(OUTPUT_EXCEL, index=False)

        if OUTPUT_EXCEL.is_file() and OUTPUT_EXCEL.stat().st_size > 0:
            print(f'✓ Excel-Datei gespeichert unter: {OUTPUT_EXCEL}')
        else:
            print('✗ Export schien zu laufen, aber die Datei ist nicht entstanden oder leer.')
    except PermissionError:
        print(f'✗ Konnte die Datei nicht schreiben (PermissionError). Ist {OUTPUT_EXCEL} evtl. geöffnet?')
    except Exception as e:
        print(f'✗ Unerwarteter Fehler beim Schreiben der Excel-Datei: {e}')

if __name__ == '__main__':
    main()
