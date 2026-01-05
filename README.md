# StreamSafe CD4+ FOXP3+ Excel Processor

A memory-efficient Python tool for processing large immunology Excel datasets. This script identifies CD4+ FOXP3+ phenotype rows, filters them, categorizes by distance, and organizes results by sample subtypes — all while staying fully streaming-safe for large workbooks.

---

## Features

- **Streaming-safe** processing using `openpyxl` (`read_only` + `write_only`).
- Automatically detects **CD4+ FOXP3+ phenotype labels** (case-insensitive).
- Filters and splits rows based on **distance ±100 microns**.
- Creates separate sheets per **subtype** based on Sample Name.
- Handles very large Excel workbooks without loading the entire file into memory.
- Ensures **unique Excel sheet names** within the 31-character limit.

---

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/streamsafecd4_foxp3.git
    cd streamsafecd4_foxp3
    ```

2. Install dependencies (Python 3.8+ recommended):
    ```bash
    pip install openpyxl
    ```

---

## Usage

1. **Configure the script**  
   Open `streamsafecd4_foxp3.py` and set your input/output file paths and sample subtype mapping at the top of the script:

    ```python
    INPUT_FILE  = "consolidated_mIF data analysis_Nov2025_JY.xlsx"
    OUTPUT_FILE = "CD4_FOXP3_filtered_output_streamsafe_v2.xlsx"

    sample_to_subtype = {
        "8F_morph": "morpheaform",
        "10A_meta": "meta",
        "2D_inf": "infiltrative",
        "5F_micro": "micronodular",
        "8D_mix": "mixed (NI)",
        "9A_nod": "nodular",
        "4B_super": "superficial"
    }
    ```

2. **Run the script**  

    ```bash
    python streamsafe_cd4_foxp3.py
    ```

3. **Check the output**  
   The script generates an Excel workbook containing:
   - `*_within100` → CD4+ FOXP3+ rows within ±100 microns
   - `*_outside100` → CD4+ FOXP3+ rows outside ±100 microns
   - Subtype-specific sheets based on Sample Name

---

## Notes

- The script automatically detects distance columns based on patterns like `distance`, `micron`, `edge`, `process`, or `tissue`.
- If no exact CD4+ FOXP3+ labels are detected, it falls back to substring matching.
- Fully streaming-safe: avoids loading entire workbooks into memory.

---

## License

MIT License
