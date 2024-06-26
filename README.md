﻿# color-exCell documentation

The code is a Python script that formats an Excel file based on the settings defined in a `settings.txt` file. Here's a breakdown of the code:

1. Import necessary libraries: `os`, `openpyxl`, `PatternFill`, and `colour`.
2. Get the file name of the only available file in the input directory.
3. Open the Excel file using `openpyxl`.
4. Read settings from the `settings.txt` file and convert color names to ARGB hex values.
5. Print all available worksheets and prompt the user to select a worksheet.
6. Determine the range of cells to format based on the selected worksheet.
7. Apply conditional formatting to the data range based on the settings defined in the `settings.txt` file.
8. Save the workbook to the output directory.

## settings.txt 

The `settings.txt` file is a text file that contains the color for each specified cell value. Each line in the file represents a value and its corresponding color. The format of each line is as follows:

```column_name=color_name```


Where:

- `column_name` is the name of the column in the Excel file (e.g., A, B, C, etc.).
- `color_name` is the name of the color to apply to the column (e.g., red, blue, yellow, etc.).

Since color names are converted to ARGB hex values using the `colour` library, you can add or replace colors in the `settings.txt` file.



## Usage


1. Install the required libraries:

```bash
pip install -r requirements.txt
```

2. Replace `/input/YOUR_EXECL_FILE` with your desired Excel file

3. Modify `settings.txt` to define the conditional formatting settings for the Excel file

4. Run the script:

```bash
python color-exCell.py
```

5. Follow the prompts to select a worksheet and apply conditional formatting based on the settings defined in the `settings.txt` file.
