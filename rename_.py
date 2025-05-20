import os
from pathlib import Path
from tkinter import Tk, filedialog

def prompt_directory():
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Select folder containing files to rename")
    return Path(folder) if folder else None

def rename_files(folder_path):
    for file in folder_path.iterdir():
        if file.is_file():
            name_without_ext = file.stem
            extension = file.suffix

            parts = name_without_ext.split('_', 2)  # Expecting 3 parts
            if len(parts) == 3:
                # Rearranged format: I-xxxxx_Dxxxxx_Company Name
                new_name = f"{parts[1]}_{parts[0]}_{parts[2]}{extension}"
                new_path = file.with_name(new_name)
                try:
                    file.rename(new_path)
                    print(f"Renamed: {file.name} → {new_name}")
                except Exception as e:
                    print(f"Error renaming {file.name}: {e}")
            else:
                print(f"Skipped (unexpected format): {file.name}")

def main():
    folder_path = prompt_directory()
    if not folder_path:
        print("No folder selected. Exiting.")
        return

    rename_files(folder_path)

if __name__ == "__main__":
    main()
