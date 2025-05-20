import pandas as pd
from pathlib import Path
import shutil
from tkinter import Tk, filedialog

def prompt_directory():
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Select folder containing Excel/CSV files")
    return Path(folder) if folder else None

def create_folders(base_path):
    for folder in ['MDR1', 'MDR2', 'MDR3', 'MDR4']:
        (base_path / folder).mkdir(exist_ok=True)

def move_file(file_path, destination_folder, reason=None):
    try:
        shutil.move(str(file_path), str(destination_folder / file_path.name))
        msg = f"Moved: {destination_folder.name} →  {file_path.name}"
        if reason:
            msg += f" (Matched: {reason})"
        print(msg)
    except Exception as e:
        print(f"Error moving {file_path.name}: {e}")

def should_skip_excel(file_path):
    try:
        with pd.ExcelFile(file_path, engine='openpyxl' if file_path.suffix.lower() == '.xlsx' else None) as excel:
            return len(excel.sheet_names) != 1
    except Exception as e:
        print(f"Error reading sheet count: {file_path.name} — {e}")
        return True

def read_file(file_path):
    try:
        if file_path.suffix.lower() == '.csv':
            return pd.read_csv(file_path, header=None, low_memory=False)

        if should_skip_excel(file_path):
            print(f"Skipped (multiple sheets): {file_path.name}")
            return None

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
        elif val_A1 == "ip_base_number" or val_B1 == "AV ID" or val_B1 == "Av ID":
            move_file(file_path, base_path / 'MDR4', "IP_BASE_NUMBER and AV ID")
        else:
            print(f"No match found: {file_path.name}")
    except Exception as e:
        print(f"Error processing {file_path.name}: {e}")

def main():
    base_path = prompt_directory()
    if not base_path:
        print("No folder selected. Exiting.")
        return

    create_folders(base_path)

    all_files = list(base_path.glob("*"))  # freeze file list before moving

    skipped = 0
    processed = 0

    for file_path in all_files:
        if file_path.suffix.lower() in ['.xlsx', '.xls', '.csv'] and file_path.is_file():
            # Remove apostrophe from filename
            if "'" in file_path.name:
                new_name = file_path.name.replace("'", "")
                new_path = file_path.with_name(new_name)
                try:
                    file_path.rename(new_path)
                    print(f"Renamed: {file_path.name} → {new_name}")
                    file_path = new_path  # update reference
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
