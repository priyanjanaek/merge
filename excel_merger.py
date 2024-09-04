import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to browse and select the Excel file
def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if file_path:
        file_path_var.set(file_path)

# Function to find and merge duplicate words in the selected Excel file
def merge_duplicates():
    file_path = file_path_var.get()

    if not file_path:
        messagebox.showwarning("No File Selected", "Please select an Excel file first.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Function to merge duplicate words in a string
        def merge_words(text):
            if isinstance(text, str):
                words = text.split()  # Split the string into words
                unique_words = list(dict.fromkeys(words))  # Remove duplicates while preserving order
                return ' '.join(unique_words)  # Join the words back into a string
            return text

        # Apply the merge_words function to every cell in the DataFrame
        merged_df = df.applymap(merge_words)

        # Store the merged DataFrame for exporting
        global merged_data
        merged_data = merged_df

        messagebox.showinfo("Success", "Duplicate words merged successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to export the merged data to a new Excel file
def export_file():
    if merged_data is None:
        messagebox.showwarning("No Data to Export", "Please merge duplicates before exporting.")
        return

    export_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )

    if export_path:
        try:
            merged_data.to_excel(export_path, index=False)
            messagebox.showinfo("Export Successful", f"Merged data exported to {export_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting: {e}")

# Initialize the main window
root = tk.Tk()
root.title("Excel Duplicate Word Merger")
root.geometry("500x200")

# Variable to store the file path
file_path_var = tk.StringVar()

# Global variable to hold the merged data
merged_data = None

# Create and place widgets
tk.Label(root, text="Select Excel File:").pack(pady=10)
tk.Entry(root, textvariable=file_path_var, width=50).pack(pady=10)
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)
tk.Button(root, text="Merge Duplicates", command=merge_duplicates).pack(pady=5)
tk.Button(root, text="Export", command=export_file).pack(pady=20)

# Run the main loop
root.mainloop()
