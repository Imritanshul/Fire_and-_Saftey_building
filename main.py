import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import messagebox
from PIL import ImageTk, Image

# Load the Excel sheet
excel_file = 'D:\Building_Python\Book1.xlsx'  # Replace with the path to your Excel file
df = pd.read_excel(excel_file)

# Create a Tkinter GUI
window = tk.Tk()
window.title("Select Values")

from PIL import Image

image_path = "D:\Building_Python\DRDO-logo.png"  # Replace with the path to your image file
image = Image.open(image_path)
image = image.resize((300, 200))  # Adjust the size as per your preference
image_tk = ImageTk.PhotoImage(image)

image_label = tk.Label(window, image=image_tk)
image_label.pack()

# Function to handle button click
def process_selection():
    column1_value = column1_var.get()
    column2_value = column2_var.get()
    column3_value = column3_var.get()

    # Find the matching row
    matching_row = df[(df['Building Types'] == column1_value) & (df['Sub-Category'] == column2_value) & (
                df['Height'] == column3_value)]

    # Check if a matching row is found
    if not matching_row.empty:
        # Create a new Word document
        doc = Document()

        # Add the matching row to the document
        table = doc.add_table(rows=1, cols=len(matching_row.columns))
        table.style = 'Table Grid'

        # Add the column headers
        for col_num, column_name in enumerate(matching_row.columns):
            table.cell(0, col_num).text = column_name

        # Add the row values
        for row_num, values in enumerate(matching_row.values):
            row_cells = table.add_row().cells
            for col_num, value in enumerate(values):
                row_cells[col_num].text = str(value)

        # Save the document
        doc.save('output.docx')
        messagebox.showinfo("Success", "Matching row saved in 'output.docx' file.")
    else:
        messagebox.showinfo("No Match", "No matching row found.")


def update_subcategories(*args):
    selected_building_type = column1_var.get()

    # Filter Sub-Category options based on selected Building Type
    filtered_subcategories = df[df['Building Types'] == selected_building_type]['Sub-Category'].unique()

    # Clear previous selection and update options
    column2_menu['menu'].delete(0, 'end')
    for subcategory in filtered_subcategories:
        column2_menu['menu'].add_command(label=subcategory, command=tk._setit(column2_var, subcategory))
    column2_var.set(filtered_subcategories[0])

    # Filter Height options based on selected Building Type and Sub-Category
    update_height_options()


def update_height_options(*args):
    selected_building_type = column1_var.get()
    selected_subcategory = column2_var.get()

    # Filter Height options based on selected Building Type and Sub-Category
    filtered_heights = df[
        (df['Building Types'] == selected_building_type) & (df['Sub-Category'] == selected_subcategory)]['Height'].unique()

    # Clear previous selection and update options
    column3_menu['menu'].delete(0, 'end')
    for height in filtered_heights:
        column3_menu['menu'].add_command(label=height, command=tk._setit(column3_var, height))
    column3_var.set(filtered_heights[0])


# Get unique values from each column
column1_values = df['Building Types'].unique()
column2_values = df['Sub-Category'].unique()
column3_values = df['Height'].unique()

# Create Tkinter variables for selected values
column1_var = tk.StringVar(window)
column2_var = tk.StringVar(window)
column3_var = tk.StringVar(window)

# Set default values for the variables
column1_var.set(column1_values[0])
column2_var.set(column2_values[0])
column3_var.set(column3_values[0])

# Create drop-down menus for selecting values
column1_menu = tk.OptionMenu(window, column1_var, *column1_values, command=update_subcategories)
column2_menu = tk.OptionMenu(window, column2_var, *column2_values, command=update_height_options)
column3_menu = tk.OptionMenu(window, column3_var, *column3_values)

# Create button to process selection
select_button = tk.Button(window, text="Select", command=process_selection)

column1_menu.config(width=40, pady=10)
column2_menu.config(width=40, pady=10)
column3_menu.config(width=40, pady=10)

# Create a frame to center the selection boxes vertically
frame = tk.Frame(window)
frame.pack(fill=tk.Y, pady=10)

# Positioning GUI elements
column1_menu.pack()
column2_menu.pack()
column3_menu.pack()
select_button.pack()

# Start the Tkinter event loop
window.mainloop()
