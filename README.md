# HiX Complication Miner – GUI for Free-Text Keyword Extraction

This Python application provides a graphical user interface (GUI) to mine and analyze complications or medical terms from free-text fields in HiX electronic health record (EHR) exports.

## Features
- **Excel file support**: Works with `.xlsx` files exported from HiX.
- **Keyword extraction**: Identifies user-defined medical keywords in free-text columns.
- **Whole-word or substring matching**: Optional setting to match keywords exactly or as substrings.
- **Stopword filtering**: Dutch stopwords are removed from the word cloud generation.
- **Dynamic results table**: Displays extracted keyword occurrences per patient.
- **Word cloud generation**: Visualizes all non-stopword terms from the dataset.
- **Configurable input**: User can set patient ID column, free-text column, keywords, and output file path.

## Requirements
- Python 3.8+
- Dependencies:
  ```bash
  pip install pandas openpyxl wordcloud pillow matplotlib
  ```

## Usage
1. Run the application:
   ```bash
   python hix_complication_miner.py
   ```
2. **Choose Excel File**: Select the `.xlsx` file exported from HiX.
3. **Select Sheet**: Pick the sheet name containing the data.
4. **Set Columns**:
   - Patient ID column (e.g., `patient_id`)
   - Text column with free-text notes (e.g., `Report`)
5. **Enter Keywords**: Comma-separated list of medical keywords.
6. **(Optional)** Check *Allow substring matches* if partial matches are needed.
7. **Extract Keywords**: Generates a table of patients with keyword occurrence flags.
8. **Generate Word Cloud**: Displays the most frequent terms (excluding stopwords).
9. **Save Output**: Export keyword extraction results to `.xlsx`.

## Output
- **Excel file**: Each row corresponds to a patient, with binary flags for each keyword.
- **Word cloud image**: Visualization of the most frequent words.

## Notes
- Handle patient identifiable data according to your organization's privacy guidelines.
- This tool is tailored for Dutch-language medical free-text but can be adapted by modifying the stopword list.

## Author
Bruno Robalo, PMC 2023-2025
