import os
import pandas as pd
from tkinter import filedialog, Tk
from pathlib import Path

def get_folder_path():
    Tk().withdraw()
    return filedialog.askdirectory(title="Select Folder with Excel Files")

def sanitize_sheetname(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).rstrip()

def split_excel_sheets(file_path):
    try:
        with pd.ExcelFile(file_path) as xl:
            if len(xl.sheet_names) <= 1:
                print(f"✅ Skipped (only 1 sheet): {file_path.name}")
                return

            for sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                safe_sheet_name = sanitize_sheetname(sheet_name)
                new_filename = f"{file_path.stem}_{safe_sheet_name}.xlsx"
                new_path = file_path.parent / new_filename
                df.to_excel(new_path, index=False)
                print(f"✔ Created: {new_filename}")

        # Delete the original file after successful split
        file_path.unlink()
        #print(f"🗑 Deleted original file: {file_path.name}")

    except Exception as e:
        print(f"‼ Error processing {file_path.name}: {e}")

def main():
    folder = get_folder_path()
    if not folder:
        print("No folder selected. Exiting.")
        return

    for item in Path(folder).iterdir():
        if item.suffix.lower() in ['.xlsx', '.xls']:
            split_excel_sheets(item)

if __name__ == "__main__":
    main()
