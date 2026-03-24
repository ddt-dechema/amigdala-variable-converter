
import csv
from pathlib import Path

 # Paths are now stored in config.py
from config import *                 

try:
    with open(datei_pfad_csv, 'r', encoding='utf-8', newline='') as csv_datei:
        # --- Hier kommt der Code zum Lesen rein (siehe nächste Schritte) ---
        datareader = csv.reader(csv_datei, delimiter=';', quotechar='"')
        print(f"Datei '{datei_pfad_csv}' erfolgreich geöffnet.")
        data_headings = []
        #for row in csv.reader(csv_datei, delimiter=';', quotechar='"'):
        #print(row)
        new_yaml = open('..//output//outfile.yaml', 'w')
        for row_index, row in enumerate(datareader):
            if row_index == 0:
                data_headings = row
            else:
                
                yaml_text = ""
                #for cell_index == 0, cell in enumerate(row):
                
                for cell_index, cell in enumerate(row):
                    if cell_index == 0:
                        lineSeperator = "- "
                        cell_text = lineSeperator  + cell.replace("\n", ", ") + ":"+ "\n"
                        yaml_text += cell_text
                    else:
                        lineSeperator = "    "
                        cell_heading = data_headings[cell_index].lower().replace(" ", "_").replace("-", "")
                        if (cell_heading == "source"+ ":"):
                            lineSeperator = '  - '

                        cell_text = lineSeperator+cell_heading + ": " + cell.replace("\n", ", ") 

                        yaml_text += cell_text+ "\n"    
                #print (yaml_text) 
                #include print if yaml text should also be printed in terminal

                new_yaml.writelines(yaml_text)
                

        new_yaml.close()    
        csv_datei.close()
except FileNotFoundError:
    print(f"FEHLER: Die Datei '{datei_pfad_csv}' wurde nicht gefunden.")


# csvfile = open('', 'r')
# ‚r‘ steht für read (lesen)
