import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

# Define the file path
file_path = r"D:\Shabbir\New project.xlsx"

# Global variable for Treeview
tree = None

# Function to handle saving data to Excel
def save_to_excel():
    # Get data from entry fields
    name = entry_name.get()
    age = entry_age.get()
    city = entry_city.get()

    # Create a DataFrame with the new data
    new_data = {
        'Name': [name],
        'Age': [int(age)],
        'City': [city]
    }
    new_df = pd.DataFrame(new_data)

    # Load existing data from Excel
    try:
        existing_df = pd.read_excel(file_path, sheet_name='Sheet1')
    except FileNotFoundError:
        existing_df = pd.DataFrame()  # Create an empty DataFrame if file doesn't exist

    # Concatenate existing and new data
    updated_df = pd.concat([existing_df, new_df], ignore_index=True)

    # Save updated DataFrame back to Excel file
    updated_df.to_excel(file_path, sheet_name='Sheet1', index=False)

    # Show success message
    messagebox.showinfo("Success", "Data updated successfully!")

    # Clear entry fields
    entry_name.delete(0, tk.END)
    entry_age.delete(0, tk.END)
    entry_city.delete(0, tk.END)

    # Update the display of updated data in the GUI
    update_display(updated_df)

# Function to update display with updated data
def update_display(dataframe):
    global tree  # Declare tree as global variable

    # Clear previous display (if any)
    for widget in display_frame.winfo_children():
        widget.destroy()

    # Display updated data in a Treeview widget
    tree = ttk.Treeview(display_frame, columns=dataframe.columns.tolist(), show="headings")
    tree.pack()

    # Insert data into Treeview
    for col in dataframe.columns.tolist():
        tree.heading(col, text=col)
    for index, row in dataframe.iterrows():
        tree.insert("", "end", values=tuple(row))

    # Add Treeview selection event
    tree.bind("<ButtonRelease-1>", on_tree_select)

# Function to handle Treeview row selection
def on_tree_select(event):
    selected_item = event.widget.selection()
    selected_indices = event.widget.selection()

    # Enable delete button if rows are selected
    if selected_indices:
        delete_button.config(state=tk.NORMAL)
    else:
        delete_button.config(state=tk.DISABLED)

# Function to delete selected rows from Treeview and Excel file
def delete_selected_rows():
    global tree  # Access the global tree variable

    selected_items = tree.selection()

    if selected_items:
        # Get indices of selected rows
        selected_indices = [tree.index(item) for item in selected_items]

        # Load existing data from Excel
        try:
            existing_df = pd.read_excel(file_path, sheet_name='Sheet1')
        except FileNotFoundError:
            existing_df = pd.DataFrame()  # Create an empty DataFrame if file doesn't exist

        # Drop selected rows from DataFrame
        updated_df = existing_df.drop(existing_df.index[selected_indices])

        # Save updated DataFrame back to Excel file
        updated_df.to_excel(file_path, sheet_name='Sheet1', index=False)

        # Show success message
        messagebox.showinfo("Success", "Selected rows deleted successfully!")

        # Update the display of updated data in the GUI
        update_display(updated_df)

# GUI Setup
root = tk.Tk()
root.title("Data Entry and Update")

# Labels and entry fields
tk.Label(root, text="Name:").pack()
entry_name = tk.Entry(root)
entry_name.pack()

tk.Label(root, text="Age:").pack()
entry_age = tk.Entry(root)
entry_age.pack()

tk.Label(root, text="City:").pack()
entry_city = tk.Entry(root)
entry_city.pack()

# Frame to display updated data
display_frame = tk.Frame(root)
display_frame.pack(pady=10)

# Button to save data
tk.Button(root, text="Save to Excel", command=save_to_excel).pack()

# Button to delete selected rows
delete_button = tk.Button(root, text="Delete Selected Rows", command=delete_selected_rows, state=tk.DISABLED)
delete_button.pack()

# Function to load existing data from Excel and update entry fields
def load_existing_data():
    try:
        existing_df = pd.read_excel(file_path, sheet_name='Sheet1')
        update_display(existing_df)
    except FileNotFoundError:
        messagebox.showwarning("File Not Found", "Excel file not found. Starting with empty fields.")

# Load existing data when GUI starts
load_existing_data()

# Run the main GUI loop
root.mainloop()
