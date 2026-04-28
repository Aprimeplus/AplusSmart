import os
import datetime

target_date = datetime.datetime(2026, 2, 1)
search_dir = r"c:\Users\Nitro V15\Desktop\AplusSmart"
output_file = r"c:\Users\Nitro V15\Desktop\AplusSmart\recent_changes.txt"

with open(output_file, "w", encoding="utf-8") as f:
    f.write(f"Scanning for files modified since {target_date}\n")
    found = False
    for root, dirs, files in os.walk(search_dir):
        if ".git" in root or "__pycache__" in root or ".venv" in root:
            continue
        for file in files:
            path = os.path.join(root, file)
            try:
                mtime = os.path.getmtime(path)
                mod_date = datetime.datetime.fromtimestamp(mtime)
                if mod_date >= target_date:
                    f.write(f"{mod_date} - {path}\n")
                    found = True
            except Exception as e:
                f.write(f"Error reading {path}: {e}\n")
    if not found:
        f.write("No files modified since target date found.\n")

print("Scan complete.")
