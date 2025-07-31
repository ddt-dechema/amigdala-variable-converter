import os
import pandas as pd
import openpyxl
import re

from pathlib import Path

# -------- Konfigurierbare Parameter --------
INPUT_DIR = 'input'  # Hauptverzeichnis mit den "x_MODELNAME"-Ordnern
OUTPUT_EXCEL = 'overview_input_files.xlsx'  # Ausgabe-Excel-Datei
# ------------------------------------------

def extract_model_name(foldername):
    """Entfernt führende Nummern und Unterstriche aus dem Ordnernamen."""
    return foldername.split('_', 1)[-1] if '_' in foldername else foldername

def get_excel_sheets_and_columns(filepath):
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

def get_csv_columns(filepath):
    """Liest die erste Zeile aus einer CSV-Datei ein, um die Spalten zu extrahieren."""
    try:
        df = pd.read_csv(filepath, nrows=1, sep=None, engine='python')
        return ', '.join([str(col) for col in df.columns])
    except Exception:
        return ''

def main():
    rows = []

    for root, dirs, files in os.walk(INPUT_DIR):
        foldername = os.path.basename(root)
        if not re.match(r'^[^_]+_', foldername):
            continue  # Nur x_MODELNAME-Ordner betrachten

        model_name = extract_model_name(foldername)

        for file in files:
            filepath = os.path.join(root, file)
            filename = os.path.basename(file)
            extension = Path(file).suffix.lower()

            if extension in ['.xls', '.xlsx']:
                sheets = get_excel_sheets_and_columns(filepath)
                for sheet, columns in sheets:
                    rows.append({
                        'Source model': model_name,
                        'File location': foldername,
                        'File name': filename,
                        'Sheet name': sheet,
                        'Column names': columns
                    })
            elif extension == '.csv':
                columns = get_csv_columns(filepath)
                rows.append({
                    'Source model': model_name,
                    'File location': foldername,
                    'File name': filename,
                    'Sheet name': '',
                    'Column names': columns
                })
            else:
                # Andere Dateitypen ignorieren
                continue

    # Export als Excel-Datei
    df = pd.DataFrame(rows)
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f'✓ Excel-Datei gespeichert unter: {OUTPUT_EXCEL}')

if __name__ == '__main__':
    main()
