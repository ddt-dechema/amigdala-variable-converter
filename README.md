# README for pyam Converter Project

## Project Overview

This project automates the conversion of modeling results (CSV/Excel) from various energy system models into the pyam format. The goal is to standardize and efficiently prepare data for a data explorer or further analysis.

## Project Structure

```
amigdala-variable-converter/
│
├── konverter/
│   ├── __init__.py
│   ├── prism.py                # Conversion logic for the PRISM model (example)
│   └── ...                     # Other model converters
│
├── input/                      # Input files (CSV/Excel, structured by model)
│   └── ...
│
├── output/                     # Output files in pyam format
│   └── ...
│
├── variable_mapping.xlsx       # Central mapping file for all models and variables
├── requirements.txt            # Dependencies
└── README.md                   # This file
```

## Setup

1. **Create a virtual environment**

   ```bash
   python -m venv venv
   # Activate (Windows)
   venv\Scripts\activate
   # Activate (Mac/Linux)
   source venv/bin/activate
   ```

2. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   # or manually:
   pip install pandas openpyxl
   ```

## Usage

1. **Maintain the mapping file**
   - The central mapping file (`variable_mapping.xlsx`) contains the mapping of original variables to pyam variables, units, definitions, etc., for all models.
   - This file can be edited with Excel and should include:
     - Source model
     - File location
     - File name
     - Variable column (if the variable is a value in a column)
     - Variable value (original)
     - Variable name (new)
     - Unit
     - Separator (if CSVs should be processed, too)
     - ... (other metadata as needed)

2. **Check for files to convert**
   - First, the python script should be run which searches for all files to be converted.
      ```bash
        konverter\lookup_files.py`
      ```
   - This script looks in every sub-folder of `\input` for `.xslx` and `.csv`-files
   - Each model runs should be saved in a subfolder written like this: `1_MODELNAME`.
   - It lists the information in `overview_input_files.xlsx` in nearly the format needed for the variables mapping-Excel file. There are:
      - model name
      - folder name
      - file name
      - sheet name
      - headers of the columns, which should contain information about: variable, region, year, unit, value
         Please note, that some model output files contain multiple sheets, although not all of them are relevant
   - These information should then be put into the `variable_mapping.xlsx` file.

3. **Run the conversion script the first time for the variables**
   - Run the conversion script with:
     ```bash
     python konverter/variable_converter.py
     ```
   - The script reads the input file(s), uses the mapping file, and generates a pyam-compatible Excel file in the `output/` folder.
   - As we want to harmonize the variable names, the mapping file contains information on the current variable name and to desired variable name.
   - The first time this script runs, it will find that all the variables have not been assigned to a new name.
   - We then use the terminal from the python script and fill copy paste them into the `variable_mapping.xlsx` file and assign them to the desired variable name.
   
4. **Run the conversion script the first time for the regions**
   - While running the conversion script the first time, it will also identify region names, which have not been mapped yet.
   - These regions should be copied to the "region_mapping" sheet of the `variable_mapping.xlsx` file.

4b. **Run the script to harmonize the region naming**
   - If necessary and some regions are named differently, e.g. with 2 letter codes or with strange 3 letter codes, use the following script to harmonize them:
      ```bash
      python konverter/convert_regions_to_iso.py
      ```
   - This script will convert the regions in the  `variable_mapping.xlsx` file in "region_mapping" to a newly created csv file `variable_mapping.xlsx` file.
   - The csv can be copied into the Excel file which will then be used by the conversion script for the renaming of the variables.


5. **Run the conversion script the second time**
   - When running the conversion script again, it will now hopefully convert the input files from the model runs into the pyam format and rename the variables and regions to the harmonized ones.
   - It is always helpful to look at the logs in the terminal, if the script encounters any errors.
   - Usually, it helps to contact the modelers and ask for their help. The most typical errors are:
      - The modelers did not include a variable at all in the resulting excel.
      - a column for the unit is missing
      - the variables are abbreviated and it is not clear, what the abbreviation stands for.
   - The resulting files are then found in the `output/` folder.
   - All files are named `pyam_MODELNAME_original-filename.xlsx`.

6. **Check the result**
   - The output files should be checked and could be uploaded directly to the data explorer or used for further analysis.

## Notes

- The mapping file is the central place for all variable, unit, and metadata harmonization. Changes are made here and immediately reflected in the conversion.
- The script is designed to handle both cases:
  - Variables as columns (wide format)
  - Variables as values in a column (long format, e.g., "sortingstream")
- Each model can have its own conversion script, or you can generalize the logic for batch processing.
- The project structure is designed to be easily extendable (add more models, more mappings, etc.).

## Next Steps

- Integrate additional model files (which were not shared on the TNO Sharepoint) by adding them to the mapping file and creating new scripts if needed.
- Check the variable lists and update the amigdala-workflow repo and inform IIASA
- upload the data to the [data explorer](https://amigdala-internal.apps.ece.iiasa.ac.at/)


---

## Example: Final pyam/IAMC File Structure

The resulting Excel/CSV file will have the following structure:

![Example pyam/IAMC table structure](iamc_template.webp)

Each row represents a unique combination of model, scenario, region, variable, and unit, with yearly data as columns:

| model | scenario | region | variable | unit | 2020 | 2025 | 2030 | ... |
|-------|----------|--------|----------|------|------|------|------|-----|
| PRISM | 1. W2.4-EU net0 | AUT | Recycling | t | 247411.99 | 54031.48 | 55913.79 | ... |


## r