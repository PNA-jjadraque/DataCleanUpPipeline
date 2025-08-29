import os
import pandas as pd
from tkinter import filedialog, Tk

# Prompt user to select folder
def get_folder_path():
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder with Excel and CSV Files")

# Clean up Columns E and F (index 4 and 5)
def clean_columns_e_f(df):
    for col_index in [4, 5]:  # Column E = 4, Column F = 5
        if df.shape[1] > col_index:
            col = df.columns[col_index]
            df[col] = (
                df[col].astype(str)
                .replace({
                    r'[\t/]|\\t|"\t"| \t|\t |	 ': ' ',  # Replace tab and slashes
                }, regex=True)
                .str.replace(' +', ' ', regex=True)  # Collapse multiple spaces
                .str.strip()
                .fillna('')
            )
    return df

# Save DataFrame as tab-separated TXT in same folder
def save_as_txt(df, folder_path, original_filename):
    txt_filename = os.path.splitext(original_filename)[0] + '.txt'
    txt_path = os.path.join(folder_path, txt_filename)
    try:
        df.to_csv(txt_path, sep='\t', index=False)
        print(f"✔ Saved: {txt_path}")
    except Exception as e:
        print(f"❌ Failed to save {txt_path}: {e}")

# Process a single file
def process_file(filepath, filename, folder_path):
    try:
        ext = filename.lower()
        if ext.endswith('.csv'):
            df = pd.read_csv(filepath)
        elif ext.endswith(('.xls', '.xlsx')):
            with pd.ExcelFile(filepath) as xl:
                if len(xl.sheet_names) != 1:
                    print(f"⚠ Skipping multi-sheet file: {filename}")
                    return
                df = xl.parse(xl.sheet_names[0])
        else:
            return

        # Drop first column if inside MDR1 or MDR2
        if any(x in filepath for x in ['MDR1', 'MDR2']) and df.shape[1] >= 1:
            df.drop(df.columns[0], axis=1, inplace=True)
            print(f"🗑 Deleted first column for: {filename}")

        df = clean_columns_e_f(df)
        save_as_txt(df, folder_path, filename)

        # ✅ Delete original file after processing
        os.remove(filepath)
        print(f"🗑 Deleted original: {filename}")

    except Exception as e:
        print(f"‼ Error with file {filename}: {e}")

# Main logic
def main():
    folder_path = get_folder_path()
    if not folder_path:
        print("No folder selected.")
        return

    for root, _, files in os.walk(folder_path):
        print(f"📂 Processing folder: {root}")
        for filename in files:
            if filename.lower().endswith(('.xls', '.xlsx', '.csv')):
                old_filepath = os.path.join(root, filename)

                # Remove apostrophe from filename
                if "'" in filename:
                    sanitized_name = filename.replace("'", "")
                    new_filepath = os.path.join(root, sanitized_name)
                    try:
                        os.rename(old_filepath, new_filepath)
                        print(f"✏ Renamed: {filename} → {sanitized_name}")
                        filename = sanitized_name
                        filepath = new_filepath
                    except Exception as e:
                        print(f"‼ Error renaming {filename}: {e}")
                        continue
                else:
                    filepath = old_filepath

                process_file(filepath, filename, root)

if __name__ == "__main__":
    main()
