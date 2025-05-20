import pandas as pd
import shutil
from pathlib import Path
from tkinter import Tk, filedialog
import time

def prompt_directory():
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Select folder containing Excel/CSV files")
    return Path(folder) if folder else None

def create_folders(base_path):
    for folder in ['MDR1', 'MDR2', 'MDR3', 'MDR4']:
        (base_path / folder).mkdir(exist_ok=True)

def sanitize_sheetname(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).rstrip()

def split_excel_sheets(file_path):
    new_files = []
    try:
        for _ in range(3):  # Retry up to 3 times
            try:
                with pd.ExcelFile(file_path, engine='openpyxl' if file_path.suffix.lower() == '.xlsx' else None) as xl:
                    if len(xl.sheet_names) <= 1:
                        return []  # nothing to split

                    for sheet_name in xl.sheet_names:
                        df = xl.parse(sheet_name)
                        safe_sheet_name = sanitize_sheetname(sheet_name)
                        new_filename = f"{file_path.stem}_{safe_sheet_name}.xlsx"
                        new_path = file_path.parent / new_filename
                        df.to_excel(new_path, index=False)
                        print(f"✔ Created: {new_filename}")
                        new_files.append(new_path)

                # Try deleting the original file
                file_path.unlink()
                print(f"🗑 Deleted original file: {file_path.name}")
                return new_files

            except PermissionError as e:
                print(f"⏳ File in use, retrying: {file_path.name}")
                time.sleep(2)
        else:
            print(f"‼ Could not access file after retries: {file_path.name}")

    except Exception as e:
        print(f"‼ Error splitting {file_path.name}: {e}")
    return new_files

def move_file(file_path, destination_folder, reason=None):
    try:
        shutil.move(str(file_path), str(destination_folder / file_path.name))
        msg = f"Moved: {destination_folder.name} → {file_path.name}"
        if reason:
            msg += f" (Matched: {reason})"
        print(msg)
    except Exception as e:
        print(f"Error moving {file_path.name}: {e}")

def read_file(file_path):
    try:
        if file_path.suffix.lower() == '.csv':
            return pd.read_csv(file_path, header=None, low_memory=False)
        return pd.read_excel(file_path, sheet_name=0, header=None,
                             engine='openpyxl' if file_path.suffix.lower() == '.xlsx' else None)
    except Exception as e:
        print(f"Error reading file: {file_path.name} — {e}")
        return None

def classify_and_move(file_path, df, base_path):
    try:
        val_I13 = df.iat[12, 8] if df.shape[0] > 12 and df.shape[1] > 8 else None
        val_A1 = df.iat[0, 0] if df.shape[0] > 0 and df.shape[1] > 0 else None
        val_B1 = df.iat[0, 1] if df.shape[0] > 0 and df.shape[1] > 1 else None

        if val_I13 == "Controlled Client Name":
            move_file(file_path, base_path / 'MDR2', "Controlled Client Name in I13")
        elif val_I13 == "Date from":
            move_file(file_path, base_path / 'MDR1', "Date from in I13")
        elif val_A1 == "ip_base_number" and val_B1 == "Distribution Pool Code":
            move_file(file_path, base_path / 'MDR3', "IP_BASE_NUMBER and Distribution Pool Code")
        elif val_A1 == "ip_base_number" and val_B1 in ["AV ID", "Av ID"]:
            move_file(file_path, base_path / 'MDR4', "IP_BASE_NUMBER and AV ID")
        else:
            print(f"No match found: {file_path.name}")
    except Exception as e:
        print(f"Error classifying {file_path.name}: {e}")

def main():
    base_path = prompt_directory()
    if not base_path:
        print("No folder selected. Exiting.")
        return

    create_folders(base_path)

    all_files = list(base_path.glob("*"))
    processed = 0
    skipped = 0

    # Step 1: Split files with multiple sheets
    extra_files = []
    for file_path in all_files:
        if file_path.suffix.lower() in ['.xlsx', '.xls'] and file_path.is_file():
            try:
                with pd.ExcelFile(file_path, engine='openpyxl' if file_path.suffix.lower() == '.xlsx' else None) as xl:
                    if len(xl.sheet_names) > 1:
                        extra_files += split_excel_sheets(file_path)
            except Exception as e:
                print(f"Error checking sheets in {file_path.name}: {e}")

    all_files += extra_files  # Add newly created single-sheet files

    # Step 2: Classify and move files
    for file_path in all_files:
        if file_path.suffix.lower() in ['.xlsx', '.xls', '.csv'] and file_path.is_file():
            if "'" in file_path.name:
                new_name = file_path.name.replace("'", "")
                new_path = file_path.with_name(new_name)
                try:
                    file_path.rename(new_path)
                    print(f"Renamed: {file_path.name} → {new_name}")
                    file_path = new_path
                except Exception as e:
                    print(f"Error renaming {file_path.name}: {e}")
                    continue

            print(f"Checking: {file_path.name}")
            df = read_file(file_path)
            if df is not None:
                classify_and_move(file_path, df, base_path)
                processed += 1
            else:
                skipped += 1

    print(f"\nSummary:")
    print(f"Processed files: {processed}")
    print(f"Skipped files: {skipped}")

if __name__ == "__main__":
    main()
