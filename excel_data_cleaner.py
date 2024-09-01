import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import shutil

# Initialize global variables
df = None
file_path = ""

# Function to load Excel file
def load_file():
    global file_path, df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            messagebox.showinfo("File Loaded", "Excel file loaded successfully!")
            preview_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to preview data
def preview_data():
    if df is not None:
        preview_window = tk.Toplevel(root)
        preview_window.title("Data Preview")
        preview_window.geometry("600x400")
        
        tree = ttk.Treeview(preview_window)
        tree.pack(expand=True, fill='both')
        
        tree["columns"] = list(df.columns)
        tree["show"] = "headings"
        
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="w")
        
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))
    else:
        messagebox.showwarning("No File", "Please load an Excel file first.")

# Function to remove duplicates
def remove_duplicates():
    global df
    if df is not None:
        cols = df.columns.tolist()
        selected_cols = []
        
        def select_columns():
            nonlocal selected_cols
            selected_cols = [cols[i] for i in col_list.curselection()]
            remove_duplicates_action(selected_cols)
            select_window.destroy()

        select_window = tk.Toplevel(root)
        select_window.title("Select Columns to Remove Duplicates")
        
        col_list = tk.Listbox(select_window, selectmode='multiple', selectbackground="#f0f0f0")
        col_list.pack(pady=10)
        for col in cols:
            col_list.insert(tk.END, col)

        tk.Button(select_window, text="Select", command=select_columns).pack(pady=5)

    else:
        messagebox.showwarning("No File", "Please load an Excel file first.")

def remove_duplicates_action(selected_cols):
    global df
    if selected_cols:
        df.drop_duplicates(subset=selected_cols, inplace=True)
        messagebox.showinfo("Duplicates Removed", f"Duplicates removed based on columns: {', '.join(selected_cols)}")
        preview_data()

# Function to save the cleaned Excel file
def save_file():
    if df is not None:
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("File Saved", "Cleaned data saved successfully!")
    else:
        messagebox.showwarning("No File", "Please load an Excel file first.")

# Function to copy the Excel file to another location
def copy_file():
    if file_path:
        destination = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if destination:
            try:
                shutil.copy(file_path, destination)
                messagebox.showinfo("File Copied", "Excel file copied successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy file: {e}")
    else:
        messagebox.showwarning("No File", "Please load an Excel file first.")

# Setting up the GUI
root = tk.Tk()
root.title("Excel Data Cleaner and Copier")
root.geometry("400x300")
root.configure(bg="#f0f0f0")

# Create a frame for better layout
frame = tk.Frame(root, bg="#f0f0f0")
frame.pack(pady=20)

# Buttons for various operations with improved styling
load_button = tk.Button(frame, text="Load Excel File", command=load_file, width=20, bg="#4CAF50", fg="white")
remove_duplicates_button = tk.Button(frame, text="Remove Duplicates", command=remove_duplicates, width=20, bg="#2196F3", fg="white")
save_button = tk.Button(frame, text="Save Cleaned File", command=save_file, width=20, bg="#FF9800", fg="white")
copy_button = tk.Button(frame, text="Copy Excel File", command=copy_file, width=20, bg="#f44336", fg="white")

# Placing the buttons in a vertical layout
load_button.pack(pady=5)
remove_duplicates_button.pack(pady=5)
save_button.pack(pady=5)
copy_button.pack(pady=5)

# Run the GUI loop
root.mainloop()
