import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import analyze as anl
import config as cfg
import process as prc


def main():
    """Data extraction utilities with visual progress bar."""
    # Initialize Tkinter and hide the main window for folder selection
    root = tk.Tk()
    root.title("Phishing Campaign Data Extraction")
    root.withdraw()

    # Windows file dialog to select the input directory containing CSV files
    input_dir = filedialog.askdirectory(
        title="Select the folder containing the phishing campaign CSV files"
    )

    if not input_dir:
        messagebox.showwarning("Operation Cancelled", "No folder was selected.")
        return

    input_dir = os.path.abspath(input_dir)
    output_dir = f"{input_dir}_results"
    os.makedirs(output_dir, exist_ok=True)
    output_file_path = os.path.join(output_dir, "phishing_results.xlsx")

    # Create a list of CSV files in the selected directory to determine the total number of files for the progress bar
    all_files = os.listdir(input_dir)
    csv_files = [
        f
        for f in all_files
        if os.path.isfile(os.path.join(input_dir, f)) and f.lower().endswith(".csv")
    ]
    total_files = len(csv_files)

    if total_files == 0:
        messagebox.showinfo("Warning", "No CSV files found in the selected folder.")
        return

    # Configure the progress window
    root.deiconify()

    width = 450
    height = 150

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate X and Y position to center the window on the screen
    position_x = (screen_width // 2) - (width // 2)
    position_y = (screen_height // 2) - (height // 2)

    # Define the size and position (WidthxHeight+X+Y) of the window
    root.geometry(f"{width}x{height}+{position_x}+{position_y}")

    root.resizable(False, False)

    # Create the progress bar and status label
    label_title = tk.Label(
        root, text="Processing phishing campaigns...", font=("Arial", 11, "bold")
    )
    label_title.pack(pady=10)

    label_status = tk.Label(root, text="Starting...", font=("Arial", 9), fg="gray")
    label_status.pack(pady=2)

    barra_progress = ttk.Progressbar(
        root, orient="horizontal", length=350, mode="determinate"
    )
    barra_progress["maximum"] = total_files
    barra_progress.pack(pady=10)

    root.update()

    # File processing phase with progress bar updates
    for index, filename in enumerate(csv_files, 1):
        input_file_path = os.path.join(input_dir, filename)

        # Update the status label and progress bar for the current file being processed
        label_status.config(text=f"File {index} of {total_files}: {filename}")
        barra_progress["value"] = index

        # Refresh the GUI to show the updated status and progress
        root.update()

        # Execute the processing function for the current file, always using the same output_file_path to append results
        prc.process_file(input_file_path, output_file_path)

    # Final phase: Generate the analytics report after all files have been processed
    label_status.config(text="Generating final analytics report...")
    barra_progress.config(mode="indeterminate")
    barra_progress.start(10)
    root.update()

    anl.generate_report(output_file_path)

    # Close the progress window and show a success message with the location of the results
    root.withdraw()
    messagebox.showinfo(
        "Success!",
        f"The analytics report was generated successfully.\n\nResults saved to:\n{output_file_path}",
    )
    root.destroy()


if __name__ == "__main__":
    cfg.banner()
    main()
