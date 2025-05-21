# SmartCare EMR Order Processor

**Read the story behind this project on Medium:** [Taming a "Smart" EMR with Code (and a Little AI Help)](https://medium.com/@pythonfordocs/taming-a-smart-emr-with-code-and-a-little-ai-help-cbccf263e59b)

This Python script is designed to help clinicians process and review medication orders exported from the SmartCare EMR system. It aims to streamline workflow, improve clarity, and aid in patient safety by converting clunky EMR exports into user-friendly, sorted Excel spreadsheets and concise text summaries.

This project was born out of the frustrating and time-consuming experiences with the SmartCare EMR, particularly its difficulties in managing and verifying complex medication orders for psychiatric patients.

## The Problem Addressed

The SmartCare EMR can be cumbersome for tasks like:
* Efficiently reviewing all active medication orders.
* Comparing current orders against notes or previous regimens.
* Quickly identifying orders nearing their end date.
* Ensuring accuracy for patients on multiple, potentially interacting psychiatric medications.

The default EMR exports are not always easy to work with, and the system's performance can lead to significant time waste and potential for error.

## Solution

This Python script automates the processing of exported SmartCare medication orders. It takes an Excel file (which you first export from SmartCare and then re-save as `.xlsx`), and then:
1.  Extracts key order information: Medication Name, Order Comments, Frequency, Start Date, and End Date.
2.  Classifies each medication order into one of four types:
    * Psychiatric Scheduled
    * Psychiatric PRN
    * Other Scheduled
    * Other PRN
3.  Sorts the orders: Primarily by the `Type` (in the order listed above), and secondarily alphabetically by medication `Name`.
4.  Generates two output files for each input file:
    * An **Excel file (`_OUT.xlsx`)** with the processed, sorted data and auto-adjusted column widths.
    * A **Text file (`_OUT.txt`)** providing a summarized list, grouped by type, with specific formatting for quick review (including conditional display of end dates within the next 7 days and abbreviated frequencies).

## Prerequisites: Exporting and Preparing Orders from SmartCare EMR

Before running the Python script, you need to export the medication orders from SmartCare and prepare the file:

1.  **Log onto SmartCare EMR.**
2.  **Search for the patient** whose orders you want to review.
3.  Navigate to the orders section and click to view **"all active orders"**.
4.  The filter for "All Types" is usually currently selected by default. Click this and change the filter to **"Medication"**.
5.  Click **"Apply Filter"**.
6.  Click the **"Export" button** (this usually downloads an `.xml` file mislabeled as `.xls`, or a similar problematic format).
7.  **Save the downloaded file** into your designated input folder for the script (e.g., `Orders_IN`). Let's call this the "original EMR download."
8.  **Crucial Step for Compatibility:** Open this "original EMR download" file using Microsoft Excel.
9.  In Excel, go to **File > Save As...**
10. From the "File Format" or "Save as type" dropdown, select **"Excel Workbook (.xlsx)"**.
11. **Save this newly converted `.xlsx` file** into the *same* input folder (e.g., `Orders_IN`). You can give it the same name as the original or a new one (e.g., `PatientName_Orders_Converted.xlsx`).

**Important:** The Python script (`SmartCare_Orders.py`) is designed to process the **re-saved `.xlsx` file**. It will likely not be able to correctly process the original direct EMR download due to its formatting issues. You do not necessarily need to delete the original EMR download from your `Orders_IN` folder, but ensure the `.xlsx` version you want to process is present.

## Setup for the Python Script

### Requirements
* Python 3 (tested with Python 3.x)
* Access to a terminal or command prompt to run the script.

### Python Libraries
The script uses several Python libraries. You can install them all with a single command using `pip3` (or `pip` if that's your system's Python 3 package installer):

```bash
pip3 install pandas openpyxl xlrd lxml xlsxwriter
````

  * **pandas:** For data manipulation and analysis.
  * **openpyxl:** For reading and writing modern `.xlsx` Excel files.
  * **xlrd:** For reading older `.xls` Excel files (used as a fallback).
  * **lxml:** For parsing XML-based Excel files (used as a fallback).
  * **xlsxwriter:** For writing `.xlsx` files with enhanced formatting, like auto-adjusted column widths.

## How to Use the Script (`SmartCare_Orders.py`)

### Folder Structure

1.  It's recommended to have a main project folder where `SmartCare_Orders.py` is located (e.g., `/Volumes/DK_DRIVE/SmartCare/Orders/`).
2.  Inside this main folder, create a subfolder named `Orders_IN` (e.g., `/Volumes/DK_DRIVE/SmartCare/Orders/Orders_IN/`). This is where you will put the `.xlsx` files you prepared from the SmartCare EMR exports.
      * *Note: The script uses the path `/Volumes/DK_DRIVE/SmartCare/Orders/Orders_IN/` as the hardcoded `input_base_dir`. If your `Orders_IN` folder is located elsewhere, you will need to adjust this variable within the `main_process_all_files` function in the script.*
3.  When you run the script from its location (e.g., from `/Volumes/DK_DRIVE/SmartCare/Orders/`), it will automatically create an `Orders_OUT` subfolder within that same directory (e.g., `/Volumes/DK_DRIVE/SmartCare/Orders/Orders_OUT/`) and save the processed files there.

### Running the Script

1.  Open your terminal or command prompt.
2.  Navigate to the directory where you saved `SmartCare_Orders.py` (e.g., `cd /Volumes/DK_DRIVE/SmartCare/Orders/`).
3.  Run the script using Python 3:
    ```bash
    python3 SmartCare_Orders.py
    ```

### Input Files

  * Place the medication order files you exported from SmartCare AND **re-saved as `.xlsx`** into the `Orders_IN` folder.
  * The script will attempt to process all `.xlsx`, `.xls`, and `.xml` files it finds in this folder, but it's optimized for the re-saved `.xlsx` files due to the EMR's original export format issues.

### Output Files

For each processed input file (e.g., `InputFile.xlsx`), the script will generate two files in the `Orders_OUT` subfolder:

1.  **`InputFile_OUT.xlsx`**: An Excel spreadsheet containing the extracted and classified orders, sorted by Type and then alphabetically by Medication Name. Column widths are auto-adjusted for better readability.
2.  **`InputFile_OUT.txt`**: A text file summary formatted for quick review:
      * Orders grouped by Type.
      * Each medication line: `Med Name with dose, Abbreviated Frequency, Order Comments, End Date (if ending within 7 days)`
      * "Abbreviated Frequency" is the content found within parentheses in the original 'Frequency' column. If no parentheses, this part will be blank in the text file.

## Features of the Output

  * **Clear Classification:** Medications are categorized into "Psychiatric Scheduled," "Psychiatric PRN," "Other Scheduled," and "Other PRN."
  * **Logical Sorting:** Ensures easy review by grouping types and alphabetizing medication names (case-insensitively).
  * **Excel Clarity:** Auto-adjusted column widths in the `.xlsx` file prevent data truncation (like `####` for dates).
  * **TXT Summary:** The `.txt` file provides a quick-glance list.
      * **Conditional End Dates:** Only displays the `End Date` if the medication order is due to end on the current day or within the next 6 days (a 7-day window from today). This helps flag orders needing imminent attention.
      * **Focused Frequency:** The text file uses only the parenthetical part of the frequency (e.g., "QHS" from "Daily at Bedtime (QHS)") for brevity.

## Important Disclaimer

This script is provided "as-is" without any warranties. It was developed to address specific challenges with one EMR system and workflow. While it aims to improve accuracy and efficiency, it is crucial to **always independently verify all clinical information and medication orders.** Clinical judgment should always supersede any output from this tool. Use at your own risk and ensure compliance with all institutional and regulatory policies.

## License

This project is licensed under the MIT License. (See the `LICENSE` file for details).

```
