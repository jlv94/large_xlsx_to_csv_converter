# Excel to CSV Converter

## Overview
This Python script converts large Excel (.xlsx) files into CSV format by reading the Excel file in manageable chunks. This approach helps avoid memory issues when working with very large Excel files.

### The script:

* Lists Excel files in a specified input folder.
* Allows the user to select which file to convert.
* Reads the Excel file in chunks (default 50,000 rows per chunk).
* Concatenates all chunks into a single DataFrame.
* Exports the combined data as a CSV file.
* Moves the original Excel file to a "done" folder after successful conversion.

### Features
* Handles large Excel files efficiently by chunked reading.
* User-friendly command-line interface.
* Configurable input, output, and archive folders.
* Uses pandas and openpyxl for Excel processing.
* Color-coded terminal messages for better user experience.

### Requirements
* Python 3.6+
* pandas
* openpyxl
* colorama

You can install the dependencies using:

``` 
pip install -r requirements.txt
```

### Usage
Prepare folders:

Create the following folders in your project directory (or specify custom paths via command-line arguments):

* tmp/ — place your .xlsx files here.
* csv/ — converted CSV files will be saved here.
* done/ — original Excel files will be moved here after conversion.


### Run the script:

```
python converter.py
```

**Select file:**
The script will list all .xlsx files in the input folder and prompt you to select one by index.

**Conversion:**
The script reads the file in chunks, concatenates the data, exports it as CSV, and moves the original Excel file to the done folder.

Example
```
python converter.py
```

### Project Structure
```
excel-to-csv-converter/
│
├── converter.py          # Main script
├── requirements.txt      # Python dependencies
├── README.md             # This file
├── tmp/                  # Input Excel files
├── csv/                  # Output CSV files
└── done/                 # Processed Excel files
```