# README for pyam Converter Project

## Project Overview

This project automates the conversion of modeling results (CSV/Excel) from various energy system models into the pyam format. The goal is to standardize and efficiently prepare data for a data explorer or further analysis.

## Project Structure

```
amigdala-variable-converter/
│
├── konverter/
│   ├── 1_lookup_files.py                # Conversion logic for the PRISM model (example)
│   └── 1b_convert_regions_to_iso.py            # Other model converters
│   └── 2_mapping_utils.py            # Other model converters
├── input/                      # Input files (CSV/Excel, structured by model)
│   └── ...                            # currently containing files from Amigdala Poc1 and Poc2 runs
│
├── output/                     # Output files in pyam format
│   └── error_log.txt            # file, which will always be created after running the mapping-utils
│   └── ...                      # all the converted pyam files
│
├── dictionary_dataexplorer_variables_translation.xlsm       # Central mapping file/dictionary for variable names, regions, scenarios, models etc. Can be found on TNO sharepoint
├── overview_files_variables.txt            # List of files and variables which are to be converted
└── README.md                   # This file
```

## Setup

1. **clone repository, install python, create virtual environment**

   ```bash
   git clone https://github.com/ddt-dechema/amigdala-variable-converter.git
   ```

   useful links:
   https://code.visualstudio.com/docs/python/python-tutorial
   

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
   pip install ...
   ```
   Do not forget to fetch:
   * model results file from TNO sharepoint
   * The central dictionary or mapping file (`dictionary_dataexplorer_variables_translation.xlsm`).
   This contains the mapping of original variables to pyam variables, as well as harmonized names for regions, scenarios etc. for all models.
   * please note that within this script, the file is renamed to "...-local.xlsm" to distinguish it from the one on the sharepoint.
   *-* The latest version of this file can be found on the TNO sharepoint

## Usage

1. **Edit the config.yaml**
   - Update the information in the config.yaml file:
      - Where are the model results stored?
      - What files should be post-processed? Please use the exemplary file as a basis.<br>
      You can use `the 1_lookup_files.py` to get an overview of files relevant for processing
      - Where is the dictionary saved at? <br>
      Here you should download and use the latest version on the sharepoint.

2. **Check for files to convert**
   - First, the python script should be run which searches for all files to be converted.
      ```bash
        python konverter\1_lookup_files.py`
      ```
   - This script looks in every sub-folder defined in the `config`-file for `.xslx` and `.csv`-files
   - if you would only like to include the files from a specific PoC run, indicate the corresponding folder
   - Each model runs should be saved in a subfolder.
   - The output of this python scripts lists the following information in the resulting file `overview_files_unsorted.xlsx`.<br>
   They are almost already in the format necessary for the variables mapping-Excel file and contain these information:
      - model name
      - folder name
      - file name
      - sheet name
      - headers of the columns.
      
      For the following processing, the files must contain information about: `variable names`, `region`, `year`, `unit`, `value`<br>
         Please note, that some model output files contain multiple sheets, although not all of them are relevant.

   >**UPDATE**
   >- This script now looks for most of the required information automatically, such as `scenario`, `region`, `year`, `value` and `unit`.
   >- Common names are stored within the code directly (lines 20-24)

   - These information should then be put into the `overview_files.xlsx` file.
   - Please note that in case of the variable names, it might be possible that they are distributed over multiple columns, each with different names.
   In that case the column names containing the variables must be concatenated with "|", e.g.:<br> 
   `Variable|Sector|Subsector|Carrier`.  

3. **Run the conversion script**
   - Run the conversion script with:
     ```bash
     python konverter/2_mapping_utils.py
     ```
   - The script reads the input file(s), uses the dictionary (mapping file), and generates a pyam-compatible Excel file in the `output/` folder for each listed excel/csv-file.
   - The first time this script runs, it might find some `variables` which are not listed in the dictionary yet, possibly also `regions`, `scenario` or `model names`.<br>
   These are listed in the terminal and the `error_log.txt` in the `output/` folder.
   - These information should be discussed bilaterally with the model owners and then updated in the dictionary file.
   - Re-run the conversion with the updated dictionary file until there are no more errors. 
   - All converted files are stored in the `/output`-folder named `pyam_MODELNAME_original-filename.xlsx`.
   > NOTE : If there are multiple files for a model, they are aggregated into a single file, because this is required for uploading to the data explorer. 

## Others
1. **Run the script to harmonize update region naming**
   - If necessary and some regions are named differently, e.g. with 2 or 3 letter codes or with other abbreviatons.<br>
   Use the following script to harmonize them:
      ```bash
      python konverter/1b_convert_regions_to_iso.py
      ```
   - This script uses the `country_converter` library to convert the regions to full names.
   Custom region names are stored directly in a python `dict`.
   - This script will convert the regions in the  `dictionary_dataexplorer_variables_translation.xlsm` file the sheet `regions` and create a new file `dictionary_dataexplorer_variables_translation-local_regions_fullname.xlsx` file.
   - The information can then be used to update the region list in the dictionary file.

## Notes

- The mapping file is the central place for all variable, unit, and metadata harmonization. Changes are made here and immediately reflected in the conversion.
- The script is designed to handle both cases:
  - Variables as columns (wide format)
  - Variables as values in a column (long format, e.g., "sortingstream")
- Each model can have its own conversion script, or you can generalize the logic for batch processing.
- The project structure is designed to be easily extendable (add more models, more mappings, etc.).

## Next Steps

- Check the variable lists and update the amigdala-workflow repo and inform IIASA via Pull request
- upload the data to the [data explorer](https://amigdala-internal.apps.ece.iiasa.ac.at/)


---

## Example: Final pyam/IAMC File Structure

The resulting Excel/CSV file will have the following structure:

![Example pyam/IAMC table structure](iamc_template.webp)

Each row represents a unique combination of model, scenario, region, variable, and unit, with yearly data as columns:

| model | scenario | region | variable | unit | 2020 | 2025 | 2030 | ... |
|-------|----------|--------|----------|------|------|------|------|-----|
| PRISM | 1. W2.4-EU net0 | AUT | Recycling | t | 247411.99 | 54031.48 | 55913.79 | ... |

---

This script was written by DDT (DECHEMA) for the AMIGDALA project.
The AMIGDALA project is funded by the European Union under the grant agreement 101138534. 