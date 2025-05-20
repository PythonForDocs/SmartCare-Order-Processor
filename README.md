# SmartCare EMR Order Processor

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
