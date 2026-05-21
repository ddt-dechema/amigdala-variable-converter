
# import csv
import pandas as pd
from pathlib import Path
from config import *

 # Paths are now stored in config.py
from config import *                 

required = ["target_yaml", "variable_name", "unit", "description"]

def norm_col(c: str) -> str:
    return str(c).strip().lower().replace(" ", "_").replace("-", "")

def clean_val(v):
    if pd.isna(v):
        return None
    s = str(v).replace("\n", ", ").strip()
    return s if s else None

def row_to_yaml(row):
    var = clean_val(row["variable_name"])
    if not var:
        return ""  # oder: raise ValueError

    lines = [f"- {var}:"]
    for col, val in row.items():
        if col in ("target_yaml", "variable_name"):
            continue
        v = clean_val(val)
        if v is None:
            continue
        lines.append(f"    {col}: {v}")
    return "\n".join(lines) + "\n\n"
        
df = pd.read_excel(datei_pfad_xlsx)  # sheet_name=sheet_name
print(f"Datei '{datei_pfad_xlsx}' erfolgreich geöffnet.")

# normalize row headings to be used as yaml keys
# (makes YAML keys predictable and makes header matching robust)
df.columns = [norm_col(c) for c in df.columns]
print(f"Spalten in Excel: {list(df.columns)}")

# check if all required columns are present in the Excel file
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError(f"Fehlende Spalten in Excel: {missing}. Vorhanden: {list(df.columns)}")

# ensure output folder is a Path (config may provide str)
OUTPUT_FOLDER = Path(OUTPUT_FOLDER)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# --- group by target_yaml, de-duplicate per (target_yaml, variable_name) ---
df = df.copy()
df["target_yaml"] = df["target_yaml"].map(clean_val)
df["variable_name"] = df["variable_name"].map(clean_val)

# drop rows missing routing/key
before = len(df)
df = df.dropna(subset=["target_yaml", "variable_name"])
after = len(df)
if after != before:
    print(f"Info: {before - after} Zeilen ohne target_yaml/variable_name wurden übersprungen.")

# de-duplicate within the Excel input: last row wins
# (prevents duplicates even if Excel is unsorted)
df = df.drop_duplicates(subset=["target_yaml", "variable_name"], keep="last")

open_files = {}  # target -> file handle


# Delete/overwrite target YAML files once per run (so reruns don't create duplicates).
target_files = list(df["target_yaml"].dropna().unique())
print("Files to be created:", target_files)
for target in target_files:
    out_path = OUTPUT_FOLDER / target
    if out_path.exists():
        out_path.unlink()  # delete file

try:
    for target, group in df.groupby("target_yaml", sort=False):
        out_path = OUTPUT_FOLDER / target

        if target not in open_files:
            # use write mode after deletion; creates a fresh file
            open_files[target] = open(out_path, "w", encoding="utf-8")

        f = open_files[target]

        for _, row in group.iterrows():
            # some columnsshould not be included in the yaml
            row_out = row.drop(labels=["variable_mapping", "source_unit"], errors="ignore")
            yaml_text = row_to_yaml(row_out)
            if yaml_text.strip():
                f.write(yaml_text)

finally:
    for f in open_files.values():
        try:
            f.close()
        except Exception:
            pass