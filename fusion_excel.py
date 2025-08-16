# Fusion Excel Tool
# Date : 17/08/2025
# Author : Doufu(DubuV2)



# ---- Import Libraries ----
import threading
import json
import pandas as pd
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, Toplevel
from tkinter import ttk

# ---- Constants ----
# Default configuration file
CONFIG_FILE = "config.json"
# ---- Configuration Management ----

def save_config(input_path, output_file, mode):
    """
    Save the configuration to a JSON file.
    Args:
        input_path (str): Path to the input folder.
        output_file (str): Path to the output file.
        mode (str): Fusion mode, either "concat" or "merge".
    """
    config = {
        "input_folder": input_path,
        "output_file": output_file,
        "mode": mode
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def load_config():
    """
    Load configuration from a JSON file and set the input, output, and mode variables.
    """
    try:
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            input_var.set(config.get("input_folder", ""))
            output_var.set(config.get("output_file", ""))
            mode_var.set(config.get("mode", "concat"))
    except FileNotFoundError:
        pass

# ---- Fusion Logic ----
def fusion_excel(input_folder: Path, output_file: Path, mode: str,):
    """
    Merge or concatenate multiple CSV or Excel files from a folder into a single file.

    Args:
        input_folder (Path): Path to the folder containing input files.
        output_file (Path): Path to the output file (CSV or Excel).
        mode (str): Fusion mode, either "concat" or "merge".
    
    Raises:
        FileNotFoundError: If the input folder does not exist or contains no CSV/Excel files.
        RuntimeError: If no files could be read successfully.
        ModuleNotFoundError: If required libraries for reading/writing files are not installed.
    """
    input_path = Path(input_folder)
    if not input_path.is_dir():
        raise FileNotFoundError(f"The folder '{input_folder}' does not exist or is not a directory.")
    
    files = list(input_path.glob('*.csv')) + \
            list(input_path.glob('*.xlsx')) + \
            list(input_path.glob('*.xls'))
    
    if not files:
        raise FileNotFoundError(f"No CSV or Excel files found in the folder '{input_folder}'.")
    

    dataframes = []
    read_errors = []

    # Init Progress Bar
    progress["maximum"] = len(files)
    root.after(0, lambda: status_label.config(text="Reading Files..."))

    # Using a loop to read each file and handle exceptions
    # Not if one file fails, the others will still be processed
    for i, file in enumerate(files, start=1):
        try:
            # Check the file extension and read accordingly
            if file.suffix.lower() == ".csv":
                try:
                    # 1. Try reading with cp1252 encoding first (common for CSV files)
                    df = pd.read_csv(file, encoding='cp1252')
                except UnicodeDecodeError:
                    try:
                        # 2. If it fails, try utf-8 encoding
                        df = pd.read_csv(file, encoding='utf-8')
                    except UnicodeDecodeError:
                        # 3. If it still fails, try latin1 encoding
                        df = pd.read_csv(file, encoding='latin1')

            else:
                # For Excel files handle their own encoding
                df = pd.read_excel(file)
            dataframes.append(df)
            
        except Exception as e:
            read_errors.append(f"{file.name}: {str(e)}")

        # Force the GUI to update the progress bar and status label
        # This is necessary to keep the GUI responsive during long operations
        # Because the GUI runs in a single thread
        root.after(0, lambda val=i: progress.config(value=val))
        root.after(0, lambda val=i: progress_label.config(text=f"{val}/{len(files)} files processed"))
    
    if not dataframes:
        details = "\n".join(read_errors) if read_errors else "No readable files."
        raise RuntimeError(f"No files could be read successfully from '{input_folder}'. Details: {details}")
    
    if read_errors:
        preview = "\n".join(read_errors[:10]) # Show only the first 10 errors
        suffix = "\n... and more errors." if len(read_errors) > 10 else ""
        root.after(0, lambda: messagebox.showwarning(
            "Some files were skipped",
            f"The following files could not be read:\n{preview}{suffix}\n\n"
        ))

    # Reset progress bar for writing phase
    progress["value"] = 0
    progress["maximum"] = 1
    root.after(0, lambda: progress_label.config(text="Merging files..."))
    root.after(0, lambda: status_label.config(text="Reading completed, starting fusion..."))

    if mode == "concat":
        fusion = pd.concat(dataframes, ignore_index=True)
    
    elif mode == "merge":
        fusion = dataframes[0]
        for df in dataframes[1:]:
            fusion = pd.merge(fusion, df, how='outer')

    root.after(0, lambda: progress_label.config(text="Fusion completed, writing to file..."))
    root.after(0, lambda: status_label.config(text=""))
    
    def show_preview(df):
        """Show a preview of the first 20 rows of the DataFrame in a new window.
        This function creates a new Toplevel window to display and avoid
        blocking the main GUI thread. It uses a Treeview widget to display
        
        Args:
            df (pd.DataFrame): The DataFrame to preview.
        """
        preview_window = Toplevel(root)
        preview_window.title("Data Preview")
        preview_window.geometry("800x400")

        tree = ttk.Treeview(preview_window)
        tree.pack(fill="both", expand=True)

        tree["columns"] = list(df.columns)
        tree["show"] = "headings"
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        for _, row in df.head(20).iterrows():
            tree.insert("", "end", values=list(row))

    show_preview(fusion)

    out_path = Path(output_file)
    try :
        if out_path.suffix.lower() == ".csv":
            fusion.to_csv(out_path, index=False)
        else:
            fusion.to_excel(out_path, index=False)

    except ModuleNotFoundError as e:
        if "openpyxl" in str(e):
            raise ModuleNotFoundError("The 'openpyxl' module is required to write Excel files. Please install it using 'pip install openpyxl'.") from e
        else:
            raise
    except Exception as e:
        raise RuntimeError(f"An error occurred while writing to '{output_file}': {e}") from e

    root.after(0, lambda: progress.config(value=1))
    root.after(0, lambda: progress_label.config(text="Done ✓"))
    root.after(0, lambda: status_label.config(text="Writing completed"))
    messagebox.showinfo("Success", f"Fusion completed successfully. Output saved to '{output_file}'.")




# ---- GUI ----
def browse_input():
    folder = filedialog.askdirectory(title="Select Input Folder")
    if folder:
        input_var.set(folder)
    
def browse_output():
    file = filedialog.asksaveasfilename(
        title="Save Output File", 
        defaultextension=".xlsx",
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
    )
    if file:
        output_var.set(file)

def run_fusion():
    # Disable the start button to prevent multiple clicks
    button_start.config(state="disabled")

    # Start the fusion process in a separate thread
    # This allows the GUI to remain responsive
    # During long-running operations 
    threading.Thread(
        target=lambda: safe_fusion(),
        daemon=True
    ).start()

def safe_fusion():
    try:
        fusion_excel(input_var.get(), output_var.get(), mode_var.get())
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("Error", str(err)))
    finally:
        save_config(input_var.get(), output_var.get(), mode_var.get())
        # Re-enable the start button after processing
        root.after(0, lambda: button_start.config(state="normal"))

root = tk.Tk()
root.title("Fusion Excel Tool")
root.resizable(False, False)
try:
    root.iconbitmap("logo.ico")
except Exception:
    pass
root.option_add("*Font", ("Segoe UI", 10))

# Variables for input, output
input_var = tk.StringVar()
output_var = tk.StringVar()
mode_var = tk.StringVar(value="concat")

load_config()


# Progress bar
progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=300)
progress.grid(row=4, column=0, columnspan=3, pady=10)
# Progress bar label
progress_label = tk.Label(root, text="0/0 files processed")
progress_label.grid(row=5, column=0, columnspan=3)
status_label = tk.Label(root, text="Waiting...")
status_label.grid(row=6, column=0, columnspan=3, pady=(0, 10))

# Input folder selection
label_input = tk.Label(root, text="Input folder:")
label_input.grid(row=0, column=0, padx=5, pady=5, sticky="w")

entry_input = tk.Entry(root, textvariable=input_var, width=40)
entry_input.grid(row=0, column=1, padx=5, pady=5)

button_browse_input = tk.Button(root, text="Browse", command=browse_input, font=("Segoe UI", 10), bg="#E6E6E6")
button_browse_input.grid(row=0, column=2, padx=5, pady=5)

# Output file selection
label_output = tk.Label(root, text="Output file:")
label_output.grid(row=1, column=0, padx=5, pady=5, sticky="w")

entry_output = tk.Entry(root, textvariable=output_var, width=40)
entry_output.grid(row=1, column=1, padx=5, pady=5)

button_browse_output = tk.Button(root, text="Save As", command=browse_output, font=("Segoe UI", 10), bg="#E6E6E6")
button_browse_output.grid(row=1, column=2, padx=5, pady=5)

# Fusion mode selection
label_mode = tk.Label(root, text="Fusion mode:")
label_mode.grid(row=2, column=0, padx=5, pady=5, sticky="w")

mode_frame = tk.Frame(root)
mode_frame.grid(row=2, column=1, columnspan=2, sticky="w")

radio_concat = tk.Radiobutton(mode_frame, text="Concat (stack rows)", variable=mode_var, value="concat")
radio_concat.grid(row=0, column=0, sticky="w", padx=5)


radio_merge = tk.Radiobutton(mode_frame, text="Merge (join columns)", variable=mode_var, value="merge")
radio_merge.grid(row=0, column=1, sticky="w", padx=5)

# Start button
button_start = tk.Button(root, text="Start", command=run_fusion, font=("Segoe UI", 10, "bold"), bg="#E6E6E6")
button_start.grid(row=3, column=0, columnspan=3, pady=10)

root.mainloop()