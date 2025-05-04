import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import statistics
file_path_global = ""

def select_file():
    # Create a Tkinter root window (hidden)
    global file_path_global  # Declare that we are using the global variable
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open the file picker dialog
    file_path = filedialog.askopenfilename(
        title="Select a File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    # Print the selected file path
    if file_path:
        file_path_global = file_path
        print(f"Selected file: {file_path}")
    else:
        print("No file selected")

def get_grade_count():
    root = tk.Tk()
    root.withdraw()

    try:
        grade_count = simpledialog.askinteger(
            "Input", "Enter the number of top grades to average:",
            minvalue=1, maxvalue=1000
        )
        if grade_count is None:
            print("No input provided, exiting.")
            exit()
        return grade_count
    except Exception as e:
        print(f"Error: {e}")
        exit()

# Call the function to open the file picker

select_file()

amount_of_grades = get_grade_count()


output_path = "updated_file.xlsx"

excel_data = pd.ExcelFile(file_path_global, engine='openpyxl')
modified_sheets = {}

for sheet_name in excel_data.sheet_names:
    # Load the sheet and handle missing values
    data = pd.read_excel(file_path_global, sheet_name=sheet_name, engine='openpyxl')
    data.fillna(0, inplace=True)

    # Identify numeric columns (excluding non-numeric like student names)
    numeric_cols = data.select_dtypes(include=['number']).columns.tolist()


    # Define the function to calculate top 15 average using pre-identified numeric columns
    def calculate_top__average(row):
        new_list = row.tolist()[1:] # everything except the name
        new_list.sort(reverse=True)
        top_grades = new_list[:amount_of_grades]
        print(statistics.mean(top_grades))
        return statistics.mean(top_grades)


    # Apply the function to all rows and add the new column
    data["final grade"] = data.apply(calculate_top__average, axis=1)

    modified_sheets[sheet_name] = data

# Save all modified sheets
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet_name, modified_data in modified_sheets.items():
        modified_data.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Updated Excel file saved as {output_path}")