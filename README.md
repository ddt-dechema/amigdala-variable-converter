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
     - ... (other metadata as needed)

2. **Run the conversion script**
   - Example for the PRISM model:
     ```bash
     python konverter/prism.py
     ```
   - The script reads the input file(s), uses the mapping file, and generates a pyam-compatible Excel file in the `output/` folder.

3. **Check the result**
   - The output file can be uploaded directly to the data explorer or used for further analysis.

## Notes

- The mapping file is the central place for all variable, unit, and metadata harmonization. Changes are made here and immediately reflected in the conversion.
- The script is designed to handle both cases:
  - Variables as columns (wide format)
  - Variables as values in a column (long format, e.g., "sortingstream")
- Each model can have its own conversion script, or you can generalize the logic for batch processing.
- The project structure is designed to be easily extendable (add more models, more mappings, etc.).

## Next Steps

- Integrate additional models by adding them to the mapping file and creating new scripts if needed.
- For questions about usage or extension, contact the development team or project lead.

---

**Happy converting and analyzing!**

---

## Example: Final pyam/IAMC File Structure

The resulting Excel/CSV file will have the following structure:

![Example pyam/IAMC table structure](images/pyam_iamc_structure.webp)

Each row represents a unique combination of model, scenario, region, variable, and unit, with yearly data as columns:

| model | scenario | region | variable | unit | 2020 | 2025 | 2030 | ... |
|-------|----------|--------|----------|------|------|------|------|-----|
| PRISM | 1. W2.4-EU net0 | AUT | Recycling | t | 247411.99 | 54031.48 | 55913.79 | ... |
