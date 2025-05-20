import pandas as pd
import os
import datetime # For date calculations
import re # Added for regular expressions (frequency abbreviation)

def classify_and_extract_order_data(input_filename_to_process, base_input_dir):
    full_input_path = os.path.join(base_input_dir, input_filename_to_process)
    df = None 

    try:
        if not os.path.exists(full_input_path):
            return f"Error: File not found at {full_input_path}."

        file_ext = os.path.splitext(input_filename_to_process)[1].lower()
        print(f"Info: Attempting to read '{input_filename_to_process}' (extension: {file_ext})")

        if file_ext == '.xml': 
            print(f"Info: Reading .xml file '{input_filename_to_process}' with engine='lxml'...")
            try:
                df = pd.read_excel(full_input_path, engine='lxml')
            except Exception as e_lxml:
                print(f"Info: Reading XML file '{input_filename_to_process}' with lxml failed ({e_lxml}), trying openpyxl...")
                try:
                    df = pd.read_excel(full_input_path, engine='openpyxl')
                except Exception as e_openpyxl_xml:
                    return (f"Error: Tried reading XML file '{input_filename_to_process}' with lxml and openpyxl. "
                            f"lxml error: {e_lxml}. Openpyxl error: {e_openpyxl_xml}.")
        
        elif file_ext == '.xls': 
            print(f"Info: Reading .xls file '{input_filename_to_process}' with engine='xlrd' (fallback to lxml if it's XML)...")
            try:
                df = pd.read_excel(full_input_path, engine='xlrd')
            except Exception as xlrd_error:
                error_msg_lower = str(xlrd_error).lower()
                if "expected bof record" in error_msg_lower or "xml" in error_msg_lower or "not a cfb file" in error_msg_lower:
                    print(f"Info: xlrd failed for '{input_filename_to_process}' (likely XML-based .xls). Error: {xlrd_error}. Attempting with engine='lxml'...")
                    try:
                        df = pd.read_excel(full_input_path, engine='lxml')
                    except Exception as e_lxml_fallback:
                        return f"Error: Misnamed .xls file '{input_filename_to_process}'. xlrd failed. lxml fallback also failed. lxml Error: {e_lxml_fallback}"
                else: 
                    return f"Error: Reading .xls file '{input_filename_to_process}' with xlrd failed for non-XML reason. Error: {xlrd_error}"
        
        elif file_ext == '.xlsx': 
            print(f"Info: Reading .xlsx file '{input_filename_to_process}' with engine='openpyxl'...")
            try:
                df = pd.read_excel(full_input_path, engine='openpyxl')
            except Exception as e_openpyxl:
                return f"Error: Failed to read .xlsx file '{input_filename_to_process}' with openpyxl. Error: {e_openpyxl}"
        else:
            return f"Error: Unsupported file extension '{file_ext}' for '{input_filename_to_process}'."

        if df is None: 
             return f"Error: DataFrame is None after read attempt for '{input_filename_to_process}'. File might be empty or unreadable."

    except ImportError as e: 
        err_str = str(e).lower()
        if 'xlrd' in err_str : return f"Error: Missing 'xlrd' library. Install with: pip3 install xlrd. Original error: {e}"
        if 'openpyxl' in err_str and file_ext != '.xlsx': # openpyxl is default for xlsx, so error is likely for XML/misnamed XLS
            return f"Error: Missing 'openpyxl' for XML/misnamed XLS. Install: pip3 install openpyxl. Error: {e}"
        if 'lxml' in err_str : return f"Error: Missing 'lxml' library. Install with: pip3 install lxml. Original error: {e}"
        # If openpyxl is missing and it's an .xlsx file, pandas' read_excel might raise a more generic error
        # or an ImportError if it explicitly tries to import openpyxl.
        if 'openpyxl' in err_str and file_ext == '.xlsx':
             return f"Error: Missing 'openpyxl' for .xlsx files. Install: pip3 install openpyxl. Error: {e}"
        return f"ImportError for '{full_input_path}': {e}."
    except Exception as e: 
        return f"General error reading '{full_input_path}': {e}"

    columns_to_extract = ['Name', 'Order Comments', 'Frequency', 'Start Date', 'End Date']
    
    if df is None:
        return f"Error: Dataframe not loaded for '{input_filename_to_process}' before column extraction."

    missing_cols = [col for col in columns_to_extract if col not in df.columns]
    if missing_cols:
        return f"Error: File '{input_filename_to_process}' is missing columns: {', '.join(missing_cols)}. Available: {df.columns.tolist()}"

    df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
    extracted_df = df[columns_to_extract].copy()

    psychiatric_meds_list = [
        'aripiprazole', 'abilify', 'aristada', 'asenapine', 'saphris', 'brexpiprazole', 'rexulti',
        'cariprazine', 'vraylar', 'clozapine', 'clozaril', 'fazaclo', 'versacloz', 'iloperidone', 'fanapt',
        'lumateperone', 'caplyta', 'lurasidone', 'latuda', 'olanzapine', 'zyprexa', 
        'paliperidone', 'invega', 'quetiapine', 'seroquel', 'risperidone', 'risperdal', 'ziprasidone', 'geodon', 
        'chlorpromazine', 'thorazine', 'fluphenazine', 'prolixin', 'fluphenazine decanoate', 'haloperidol', 'haldol', 
        'haloperidol decanoate', 'loxapine', 'loxitane', 'adasuve', 'perphenazine', 'trilafon', 'pimozide', 'orap', 
        'thioridazine', 'mellaril', 'thiothixene', 'navane', 'trifluoperazine', 'stelazine', 'citalopram', 'celexa', 
        'escitalopram', 'lexapro', 'fluoxetine', 'prozac', 'sarafem', 'fluvoxamine', 'luvox', 'paroxetine', 'paxil', 
        'pexeva', 'sertraline', 'zoloft', 'desvenlafaxine', 'pristiq', 'khedezla', 'duloxetine', 'cymbalta', 'divalproex','divalproex sodium', 'depakote', 
        'levomilnacipran', 'fetzima', 'milnacipran', 'savella', 'venlafaxine', 'effexor', 'amitriptyline', 'elavil', 
        'amoxapine', 'asendin', 'clomipramine', 'anafranil', 'desipramine', 'norpramin', 'doxepin', 'sinequan', 'silenor', 
        'imipramine', 'tofranil', 'nortriptyline', 'pamelor', 'protriptyline', 'vivactil', 'trimipramine', 'surmontil', 
        'isocarboxazid', 'marplan', 'phenelzine', 'nardil', 'selegiline', 'emsam', 'tranylcypromine', 'parnate', 
        'bupropion', 'wellbutrin', 'forfivo', 'zyban', 'mirtazapine', 'remeron', 'nefazodone', 'serzone', 'trazodone', 
        'desyrel', 'oleptro', 'vilazodone', 'viibryd', 'vortioxetine', 'trintellix', 'brintellix', 'esketamine', 'spravato', 
        'alprazolam', 'xanax', 'chlordiazepoxide', 'librium', 'clonazepam', 'klonopin', 'clorazepate', 'tranxene', 
        'diazepam', 'valium', 'diastat', 'lorazepam', 'ativan', 
        'oxazepam', 'serax', 'temazepam', 'restoril', 
        'buspirone', 'buspar', 'hydroxyzine', 'vistaril', 'atarax', 'propranolol', 'inderal', 'gabapentin', 'neurontin', 
        'pregabalin', 'lyrica', 'carbamazepine', 'tegretol', 'equetro', 'valproic acid', 'depakene', 
        'lamotrigine', 'lamictal', 'lithium', 'lithobid', 'eskalith', 'oxcarbazepine', 'trileptal', 'topiramate', 'topamax', 
        'benztropine', 'cogentin', 'amphetamine', 'dextroamphetamine', 'adderall', 'mydayis', 'evekeo', 'zenzedi', 'dexedrine', 
        'dexmethylphenidate', 'focalin', 'lisdexamfetamine', 'vyvanse', 'methamphetamine', 'desoxyn', 'methylphenidate', 
        'ritalin', 'concerta', 'metadate', 'daytrana', 'quillivant', 'jornay', 'aptensio', 'cotempla', 'quillichew', 
        'atomoxetine', 'strattera', 'clonidine', 'kapvay', 'catapres', 'guanfacine', 'intuniv', 'tenex', 'viloxazine', 'qelbree', 
        'eszopiclone', 'lunesta', 'lemborexant', 'dayvigo', 'ramelteon', 'rozerem', 'suvorexant', 'belsomra', 
        'zaleplon', 'sonata', 'zolpidem', 'ambien', 'edluar', 'zolpimist', 'diphenhydramine', 'benadryl', 'unisom', 
        'doxylamine', 'unisom'
    ]
    psychiatric_meds_lower = [med.lower() for med in psychiatric_meds_list]
    
    def classify_order_row(row):
        name_str = str(row['Name']).lower()
        frequency_str = str(row['Frequency']).lower()
        is_prn = "prn" in frequency_str
        
        is_psychiatric = False
        name_parts = name_str.split()
        for part in name_parts:
            if part in psychiatric_meds_lower:
                is_psychiatric = True
                break
        if not is_psychiatric:
             for med in psychiatric_meds_lower:
                if med in name_str:
                    is_psychiatric = True
                    break
        
        if is_prn:
            if is_psychiatric:
                return "Psychiatric PRN"
            else:
                return "Other PRN"
        else: 
            if is_psychiatric:
                return "Psychiatric Scheduled"
            else:
                return "Other Scheduled"

    extracted_df['Type'] = extracted_df.apply(classify_order_row, axis=1)
    final_df = extracted_df[['Type'] + columns_to_extract] 

    type_order = [
        "Psychiatric Scheduled",
        "Psychiatric PRN",
        "Other Scheduled",
        "Other PRN"
    ]
    final_df['Type'] = pd.Categorical(final_df['Type'], categories=type_order, ordered=True)
    final_df['sort_key_name'] = final_df['Name'].astype(str).str.lower()
    final_df_sorted = final_df.sort_values(by=['Type', 'sort_key_name'])
    final_df_sorted = final_df_sorted.drop(columns=['sort_key_name'])
    
    base_name_without_ext = os.path.splitext(input_filename_to_process)[0]
    output_main_dir = os.getcwd() 
    output_sub_dir = "Orders_OUT"
    output_dir_path = os.path.join(output_main_dir, output_sub_dir)
    os.makedirs(output_dir_path, exist_ok=True) 

    excel_output_filename = f"{base_name_without_ext}_OUT.xlsx"
    excel_output_path = os.path.join(output_dir_path, excel_output_filename)
    
    # --- MODIFIED: Save to Excel with xlsxwriter for column widths ---
    try:
        with pd.ExcelWriter(excel_output_path, engine='xlsxwriter',
                            datetime_format='mm/dd/yyyy', 
                            date_format='mm/dd/yyyy') as writer:
            final_df_sorted.to_excel(writer, index=False, sheet_name='Orders')
            workbook = writer.book
            worksheet = writer.sheets['Orders']
            # Auto-adjust columns to fit content
            for i, col in enumerate(final_df_sorted.columns):
                series = final_df_sorted[col]
                max_len = max((
                    series.astype(str).map(len).max(),  # len of largest item
                    len(str(series.name))  # len of column name/header
                )) + 1  # adding a little extra space
                if col == 'Order Comments': max_len = max(max_len, 35) # Ensure comments has decent width
                if col == 'Name': max_len = max(max_len, 30) # Ensure Name has decent width
                worksheet.set_column(i, i, max_len) # Set column width
        excel_save_message = f"Successfully processed '{input_filename_to_process}'. Excel output saved to: {excel_output_path}"
    except ImportError:
        # Fallback if xlsxwriter is not installed
        try:
            final_df_sorted.to_excel(excel_output_path, index=False)
            excel_save_message = (f"Successfully processed '{input_filename_to_process}'. Excel output saved to: {excel_output_path}. "
                                  f"(Note: xlsxwriter not found, column widths not auto-adjusted. Install with 'pip3 install xlsxwriter' for better formatting.)")
        except Exception as e_fallback:
            excel_save_message = f"Error exporting '{input_filename_to_process}' to Excel (fallback): {e_fallback}"
    except Exception as e:
        excel_save_message = f"Error exporting '{input_filename_to_process}' to Excel with xlsxwriter: {e}"
    print(excel_save_message)
    # --- END MODIFIED Excel Save ---

    txt_output_filename = f"{base_name_without_ext}_OUT.txt"
    txt_output_path = os.path.join(output_dir_path, txt_output_filename)
    
    try:
        today = datetime.date.today()
        limit_date = today + datetime.timedelta(days=6) 

        df_for_txt = final_df_sorted.copy()
        # 'End Date' is already pd.to_datetime from earlier df['End Date'] conversion
        # No need to convert again if final_df_sorted is based on that.
        # However, if using a fresh copy or unsure, this is safe:
        df_for_txt['End Date'] = pd.to_datetime(df_for_txt['End Date'], errors='coerce')

        with open(txt_output_path, 'w') as f_txt:
            first_category_for_txt = True
            for type_category in type_order:
                category_df = df_for_txt[df_for_txt['Type'].astype(str) == type_category]

                if not category_df.empty or first_category_for_txt:
                    if not first_category_for_txt:
                        f_txt.write("\n") 
                    f_txt.write(f"{type_category}\n")
                    first_category_for_txt = False
                
                if category_df.empty:
                    continue 
                    
                for index, row in category_df.iterrows():
                    parts = []
                    name_dose = str(row['Name']) if pd.notna(row['Name']) else ""
                    
                    # --- MODIFIED: Abbreviated Frequency for TXT ---
                    frequency_full = str(row['Frequency']) if pd.notna(row['Frequency']) else ""
                    match = re.search(r'\((.*?)\)', frequency_full) # Non-greedy match inside parentheses
                    abbrev_freq = "" # Default to blank if no parenthesis content
                    if match:
                        abbrev_freq = match.group(1).strip() # Get content and strip spaces
                    # --- END MODIFIED Frequency for TXT ---
                    
                    parts.append(f"{name_dose}, {abbrev_freq}")

                    order_comments_val = row['Order Comments']
                    if pd.notna(order_comments_val) and str(order_comments_val).strip():
                        parts.append(str(order_comments_val).strip())
                    
                    end_date_val = row['End Date'] 
                    if pd.notna(end_date_val): 
                        end_date_date_obj = end_date_val.date()
                        if today <= end_date_date_obj <= limit_date:
                            month = str(end_date_val.month)
                            day = str(end_date_val.day)
                            year = str(end_date_val.year)
                            end_date_formatted = f"{month}/{day}/{year}"
                            parts.append(f"End Date {end_date_formatted}")
                    
                    # Join only non-empty parts, but ensure the first part (Name, Freq) is always there
                    # If abbrev_freq is blank, it will look like "Name, , Comment..."
                    # To avoid "Name, , Comment", only add abbrev_freq if it's not blank in the initial string construction.
                    line_base = str(name_dose)
                    if abbrev_freq:
                        line_base += f", {abbrev_freq}"
                    
                    # Rebuild parts list for cleaner comma handling
                    final_line_parts = [line_base]
                    if pd.notna(order_comments_val) and str(order_comments_val).strip():
                        final_line_parts.append(str(order_comments_val).strip())
                    
                    if pd.notna(end_date_val) and (today <= end_date_val.date() <= limit_date):
                        month = str(end_date_val.month)
                        day = str(end_date_val.day)
                        year = str(end_date_val.year)
                        final_line_parts.append(f"End Date {month}/{day}/{year}")

                    f_txt.write(", ".join(final_line_parts) + "\n")
        
        txt_save_message = f"TXT output saved to: {txt_output_path}"
        return f"{excel_save_message}\n{txt_save_message}"

    except Exception as e:
        txt_save_message = f"Error generating TXT file '{txt_output_path}': {e}"
        print(txt_save_message)
        return f"{excel_save_message}\n{txt_save_message}" 

def main_process_all_files():
    input_base_dir = "/Volumes/DK_DRIVE/SmartCare/Orders/Orders_IN/"

    if not os.path.isdir(input_base_dir):
        print(f"Error: Input directory not found: {input_base_dir}")
        return

    print(f"Scanning for .xml, .xlsx, and .xls files in: {input_base_dir}")
    
    files_to_process = [
        f for f in os.listdir(input_base_dir) 
        if (f.lower().endswith(".xml") or \
            f.lower().endswith(".xlsx") or \
            f.lower().endswith(".xls")) and \
           not f.startswith("~") and \
           not f.startswith("._")
    ]

    if not files_to_process:
        print(f"No processable .xml, .xlsx, or .xls files found in {input_base_dir}.")
        return

    print(f"Found {len(files_to_process)} Excel/XML file(s) to process: {', '.join(files_to_process)}")
    
    processed_files_count = 0
    error_files_count = 0

    for filename in files_to_process:
        print(f"\nProcessing file: {filename}...")
        result_message = classify_and_extract_order_data(filename, input_base_dir)
        print(result_message) 
        if "Error exporting" in result_message or ("Error reading" in result_message and "Excel output saved to" not in result_message) or "Error generating TXT" in result_message:
            error_files_count +=1
            if "Successfully processed" in result_message and "Excel output saved to" in result_message:
                 pass # Counted as error if TXT failed but Excel part was ok.
            else: # Full processing error
                 pass
        else: # No "Error" implies success for both
             processed_files_count += 1

    final_output_dir_path = os.path.join(os.getcwd(), "Orders_OUT")
    print(f"\n--- Batch Processing Summary ---")
    print(f"Total processable Excel/XML files found: {len(files_to_process)}")
    print(f"Successfully processed files (both Excel & TXT): {processed_files_count}")
    print(f"Files with errors (in Excel or TXT generation): {error_files_count}")
    print(f"Output files are saved in: {final_output_dir_path}")

if __name__ == "__main__":
    main_process_all_files()