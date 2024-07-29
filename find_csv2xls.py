#!/usr/bin/python
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback
import re
import csv
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
#from openpyxl import *
#import xlsxwriter

default_config_file = os.getcwd().replace('\\', '/') + '/new.cfg'


def search_csv(csv_file, search_query, export_folder, usecols=[]):
    # Initialize an empty list to store matching rows
    print('search_query',search_query)
    matching_rows = []

    # Use pandas chunksize parameter to read the CSV file in smaller chunks
    chunksize = 1000  # You can adjust this value based on available RAM
    if usecols:
        for chunk in pd.read_csv(csv_file, sep=',', chunksize=chunksize, quotechar='"', quoting=1, escapechar='\\', usecols=usecols):
    #    for chunk in pd.read_csv(csv_file, sep=',', chunksize=chunksize, quotechar='"', quoting=csv.QUOTE_ALL):
            # Use DataFrame query to filter rows efficiently
            query = " & ".join(f"({condition})" for condition in search_query)
            matching_chunk = chunk.query(query)
            matching_rows.append(matching_chunk)
    else:
        for chunk in pd.read_csv(csv_file, sep=',', chunksize=chunksize, quotechar='"', quoting=1, escapechar='\\'):
    #    for chunk in pd.read_csv(csv_file, sep=',', chunksize=chunksize, quotechar='"', quoting=csv.QUOTE_ALL):
            # Use DataFrame query to filter rows efficiently
            query = " & ".join(f"({condition})" for condition in search_query)
            matching_chunk = chunk.query(query)
            matching_rows.append(matching_chunk)

    # Concatenate all the matching chunks into a single DataFrame
    result_df = pd.concat(matching_rows, ignore_index=True)

    # Export the results to a new CSV file in the export folder
    if output_format_options.get() == 'csv':
        export_file = os.path.join(export_folder, f"{os.path.basename(csv_file).replace('.csv','')}_SearchResults.csv")
        result_df.to_csv(export_file, index=False, quoting=csv.QUOTE_ALL)
    else:
        export_file = os.path.join(export_folder, f"{os.path.basename(csv_file).replace('.csv','')}_SearchResults.xlsx")
        result_df.to_excel(export_file, index=False)

    print('result_df',result_df)

    return export_file, len(result_df)

def search_excel(excel_file, search_query, export_folder, usecols=[]):
    # Initialize an empty list to store matching rows
    print('search_query',search_query)
    matching_rows = []

    # Use pandas chunksize parameter to read the CSV file in smaller chunks
#    chunksize = 1000  # You can adjust this value based on available RAM
    df_excel = object()
    if usecols:
        df_excel = pd.read_excel(excel_file, usecols=usecols)
    else:
        df_excel = pd.read_excel(excel_file)

    query = " & ".join(f"({condition})" for condition in search_query)
    matching_excel = df_excel.query(query)
    matching_rows.append(matching_excel)

    # Concatenate all the matching chunks into a single DataFrame
    result_df = pd.concat(matching_rows, ignore_index=True)

    # Export the results to a new CSV file in the export folder
    if output_format_options.get() == 'csv':
        if excel_file.endswith('.xls'):
            export_file = os.path.join(export_folder, f"{os.path.basename(excel_file).replace('.xls','')}_SearchResults.csv")
        else:
            export_file = os.path.join(export_folder, f"{os.path.basename(excel_file).replace('.xlsx','')}_SearchResults.csv")
        result_df.to_csv(export_file, index=False, quoting=csv.QUOTE_ALL)
    else:
        if excel_file.endswith('.xls'):
            export_file = os.path.join(export_folder, f"{os.path.basename(excel_file).replace('.xls','')}_SearchResults.xlsx")
        else:
            export_file = os.path.join(export_folder, f"{os.path.basename(excel_file).replace('.xlsx','')}_SearchResults.xlsx")
        result_df.to_excel(export_file, index=False)

    print('result_df',result_df)

    return export_file, len(result_df)

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(tk.END, folder_path)
    populate_column_options()
    fields_entry.delete(0, tk.END)
    fields_entry.insert(tk.END, default_config_file)


def get_columns_filter(config_file):
    all_columns_list = []
    if default_config_file != config_file:
        cfg_in = open(config_file)
        for line in cfg_in.readlines():
            rline = re.search('^(.+)$', line)
            if rline:
                column = rline.group(1)
                all_columns_list.append(column)
        cfg_in.close()

    return all_columns_list


def browse_fields_filter():
    config_file = filedialog.askopenfilename(initialdir = os.getcwd(),
                                          title = "Select a Config file",
                                          filetypes = (("cfg files",
                                                        "*.cfg*"),
                                                       ("all files",
                                                        "*.*")))
    if config_file:
        fields_entry.delete(0, tk.END)
        fields_entry.insert(tk.END, config_file)

        all_columns_list = get_columns_filter(config_file)

        column1_options.set(all_columns_list[0])  # Set the first column as the default
        column2_options.set(all_columns_list[0])  # Set the first column as the default
        column3_options.set(all_columns_list[0])  # Set the first column as the default

        column1_dropdown["menu"].delete(0, tk.END)  # Clear the existing options
        column2_dropdown["menu"].delete(0, tk.END)  # Clear the existing options
        column3_dropdown["menu"].delete(0, tk.END)  # Clear the existing options

        for column in all_columns_list:
            column1_dropdown["menu"].add_command(label=column, command=tk._setit(column1_options, column))
            column2_dropdown["menu"].add_command(label=column, command=tk._setit(column2_options, column))
            column3_dropdown["menu"].add_command(label=column, command=tk._setit(column3_options, column))

def execute_search():
    folder_path = folder_entry.get()
    config_file = fields_entry.get()
    column1 = column1_options.get()
    value1 = value1_entry.get()
    column2 = column2_options.get()
    value2 = value2_entry.get()
    column3 = column3_options.get()
    value3 = value3_entry.get()

    if not folder_path or not column1 or not value1:
        messagebox.showwarning("Missing Information", "Please provide all the required information.")
        return

    # Define total_rows here
    total_rows = 0

    # Find all CSV files in the folder and its subdirectories
    csv_files = []
    excel_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".csv"):
                csv_files.append(os.path.join(root, file))
            if file.endswith(".xls") or file.endswith(".xlsx"):
                excel_files.append(os.path.join(root, file))

    if not csv_files and not excel_files:
        messagebox.showinfo("No CSV or Excel Files", "No CSV or Excel files found in the selected folder.")
        return

    for csv_file in csv_files:
        export_file = ''
        try:
            search_query = []
            if value1:
                if search_type1_var.get() == "contains":
                    search_query.append(f"`{column1}`.str.contains('{value1}', case=False, na=False)")
                else:  # "exact match"
                    search_query.append(f"{column1}.str.lower() == '{value1.lower()}'")

            if value2 and column2:
                if search_type2_var.get() == "contains":
                    search_query.append(f"`{column2}`.str.contains('{value2}', case=False, na=False)")
                else:  # "exact match"
                    search_query.append(f"{column2}.str.lower() == '{value2.lower()}'")

            if value3 and column3:
                if search_type3_var.get() == "contains":
                    search_query.append(f"`{column3}`.str.contains('{value3}', case=False, na=False)")
                else:  # "exact match"
                    search_query.append(f"{column3}.str.lower() == '{value3.lower()}'")

            columns_filter = get_columns_filter(config_file)
            export_file, rows_found = search_csv(csv_file, search_query, folder_path, columns_filter)  # Use 'folder_path' as the export folder
            total_rows += rows_found

        except Exception as e:
            traceback.print_exc()  # Print the full stack trace of the exception
            messagebox.showerror("Error", f"An error occurred while processing {csv_file}: {str(e)}")

    for excel_file in excel_files:
        export_file = ''
        try:
            search_query = []
            if value1:
                if search_type1_var.get() == "contains":
                    search_query.append(f"`{column1}`.str.contains('{value1}', case=False, na=False)")
                else:  # "exact match"
                    search_query.append(f"{column1}.str.lower() == '{value1.lower()}'")

            if value2 and column2:
                if search_type2_var.get() == "contains":
                    search_query.append(f"`{column2}`.str.contains('{value2}', case=False, na=False)")
                else:  # "exact match"
                    search_query.append(f"{column2}.str.lower() == '{value2.lower()}'")

            if value3 and column3:
                if search_type3_var.get() == "contains":
                    search_query.append(f"`{column3}`.str.contains('{value3}', case=False, na=False)")
                else:  # "exact match"
                    search_query.append(f"{column3}.str.lower() == '{value3.lower()}'")

            columns_filter = get_columns_filter(config_file)
            export_file, rows_found = search_excel(excel_file, search_query, folder_path, columns_filter)  # Use 'folder_path' as the export folder
            total_rows += rows_found

        except Exception as e:
            traceback.print_exc()  # Print the full stack trace of the exception
            messagebox.showerror("Error", f"An error occurred while processing {csv_file}: {str(e)}")

    messagebox.showinfo("Search Completed", f"Search completed for {len(csv_files)} CSV file(s).\nTotal rows exported: {total_rows}")

def populate_column_options():
    cfg_filename = os.getcwd().replace('\\', '/') + '/new.cfg'
    folder_path = folder_entry.get()
    if folder_path:
        csv_files   = []
        excel_files = []

        # Find all CSV files in the folder and its subdirectories
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".csv"):
                    csv_files.append(os.path.join(root, file))
                if file.endswith(".xls") or file.endswith(".xlsx"):
                    excel_files.append(os.path.join(root, file))

        if csv_files or excel_files:
            try:
                all_columns = set()  # Use a set to store unique column names from all CSV files
                for csv_file in csv_files:
                    try:
                        df = pd.read_csv(csv_file, encoding='utf-8', nrows=1)
                    except:
                        df = pd.read_csv(csv_file, encoding='utf-16', nrows=1)

                    columns = set(df.columns.tolist())
                    all_columns.update(columns)

                for excel_file in excel_files:
                    df = pd.read_excel(excel_file, nrows=1)

                    columns = set(df.columns.tolist())
                    all_columns.update(columns)

                all_columns_list = sorted(list(all_columns))  # Sort the columns for consistent display
                column1_options.set(all_columns_list[0])  # Set the first column as the default
                column2_options.set(all_columns_list[0])  # Set the first column as the default
                column3_options.set(all_columns_list[0])  # Set the first column as the default

                column1_dropdown["menu"].delete(0, tk.END)  # Clear the existing options
                column2_dropdown["menu"].delete(0, tk.END)  # Clear the existing options
                column3_dropdown["menu"].delete(0, tk.END)  # Clear the existing options

                for column in all_columns_list:
                    column1_dropdown["menu"].add_command(label=column, command=tk._setit(column1_options, column))
                    column2_dropdown["menu"].add_command(label=column, command=tk._setit(column2_options, column))
                    column3_dropdown["menu"].add_command(label=column, command=tk._setit(column3_options, column))

                cfg_out = open(cfg_filename, 'w')
                for column in all_columns_list:
                    cfg_out.write(column +'\n')
                cfg_out.close()

            except Exception as e:
                traceback.print_exc()  # Print the full stack trace of the exception
                messagebox.showerror("Error", f"An error occurred while reading the CSV files: {str(e)}")
        else:
            column1_dropdown["menu"].delete(0, tk.END)  # Clear the existing options
            column2_dropdown["menu"].delete(0, tk.END)  # Clear the existing options
            column3_dropdown["menu"].delete(0, tk.END)  # Clear the existing options

# Create the GUI window
window = tk.Tk()
window.title("CSV Search and Export")
window.geometry("1000x300")  # Adjust the window siz

# Create the "Search" button and associate it with the execute_search function
search_button = tk.Button(window, text="Search", command=execute_search)

# Create and place the widgets using the grid layout manager
folder_label = tk.Label(window, text="Folder:")
folder_entry = tk.Entry(window, width=60)
browse_folder_button = tk.Button(window, text="Browse", command=browse_folder)

# Create and place the widgets using the grid layout manager
fields_label = tk.Label(window, text="Fields filter:")
fields_entry = tk.Entry(window, width=60)
browse_file_button = tk.Button(window, text="Browse", command=browse_fields_filter)

column_label = tk.Label(window, text="Column:")
column1_options = tk.StringVar(window)
column1_dropdown = tk.OptionMenu(window, column1_options, "")
column2_options = tk.StringVar(window)
column2_dropdown = tk.OptionMenu(window, column2_options, "")
column3_options = tk.StringVar(window)
column3_dropdown = tk.OptionMenu(window, column3_options, "")

output_format_label = tk.Label(window, text="Output format:")
output_format_options = tk.StringVar(window)
output_format_list = ['xlsx', 'csv']
output_format_options.set(output_format_list[0])
output_format_dropdown = tk.OptionMenu(window, output_format_options, *output_format_list)

search_label1 = tk.Label(window, text="Value 1:")
value1_entry = tk.Entry(window, width=60)

search_label2 = tk.Label(window, text="Value 2:")
value2_entry = tk.Entry(window, width=60)

search_label3 = tk.Label(window, text="Value 3:")
value3_entry = tk.Entry(window, width=60)

search_type1_var = tk.StringVar(window)
search_type1_var.set("contains")
search_type1_dropdown = tk.OptionMenu(window, search_type1_var, "contains", "exact match")

search_type2_var = tk.StringVar(window)
search_type2_var.set("contains")
search_type2_dropdown = tk.OptionMenu(window, search_type2_var, "contains", "exact match")

search_type3_var = tk.StringVar(window)
search_type3_var.set("contains")
search_type3_dropdown = tk.OptionMenu(window, search_type3_var, "contains", "exact match")

# Place the widgets using the grid layout manager
folder_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
folder_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky=tk.W + tk.E)
browse_folder_button.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)

fields_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
fields_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky=tk.W + tk.E)
browse_file_button.grid(row=1, column=4, padx=5, pady=5, sticky=tk.W)

column_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
column1_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W + tk.E)
search_label1.grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
value1_entry.grid(row=2, column=3, padx=5, pady=5, sticky=tk.W + tk.E)

column2_dropdown.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W + tk.E)
search_label2.grid(row=3, column=2, padx=5, pady=5, sticky=tk.W)
value2_entry.grid(row=3, column=3, padx=5, pady=5, sticky=tk.W + tk.E)

column3_dropdown.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W + tk.E)
search_label3.grid(row=4, column=2, padx=5, pady=5, sticky=tk.W)
value3_entry.grid(row=4, column=3, padx=5, pady=5, sticky=tk.W + tk.E)

search_type1_dropdown.grid(row=2, column=4, padx=5, pady=5, sticky=tk.W)
search_type2_dropdown.grid(row=3, column=4, padx=5, pady=5, sticky=tk.W)
search_type3_dropdown.grid(row=4, column=4, padx=5, pady=5, sticky=tk.W)

output_format_label.grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
output_format_dropdown.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W + tk.E)

search_button.grid(row=6, column=0, columnspan=5, padx=5, pady=5, sticky=tk.W + tk.E)

folder_entry.bind("<FocusOut>", lambda event: populate_column_options())

def main():
    try:
        # Start the GUI event loop
        window.mainloop()

    except Exception as e:
        traceback.print_exc()  # Print the full stack trace of the exception

if __name__ == "__main__":
    main()
