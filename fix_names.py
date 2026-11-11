# fix_names.py - fixes garbled (mojibake) accented names in an Excel workbook.
# Run in IDLE: open this file, press F5.
import glob, os, shutil, sys

try:
    import openpyxl
except ImportError:
    print("openpyxl is missing. Open Command Prompt and run:  py -m pip install openpyxl")
    sys.exit()

# Known garbled -> correct pairs (add more here if the script reports any)
MAPPING = {
    "MarcAndrÃ©": "Marc-André",
    "Ã“scar": "Óscar",
    "PiquÃ©": "Piqué",
}

MOJIBAKE_SIGNS = ("Ã", "Â", "Å", "Ä", "â€")

def auto_fix(text):
    # Undo UTF-8 text that was read as Windows-1252 / Latin-1
    for enc in ("cp1252", "latin-1"):
        try:
            return text.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return None

def fix(text):
    for bad, good in MAPPING.items():
        text = text.replace(bad, good)
    if any(s in text for s in MOJIBAKE_SIGNS):
        fixed = auto_fix(text)
        if fixed:
            return fixed, True
        return text, False
    return text, True

def find_file():
    home = os.path.expanduser("~")
    files = []
    for folder in ("Downloads", "Documents", "Desktop", r"OneDrive\Desktop", r"OneDrive\Documents"):
        files += glob.glob(os.path.join(home, folder, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    return max(files, key=os.path.getmtime) if files else None

path = input("Paste the Excel file path (or press Enter to use your most recent .xlsx): ").strip().strip('"')
if not path:
    path = find_file()
if not path or not os.path.exists(path):
    print("File not found.")
    sys.exit()

print("Using:", path)
backup = path[:-5] + "_backup.xlsx"
shutil.copy(path, backup)
print("Backup saved:", backup)

wb = openpyxl.load_workbook(path)
changed, unknown = 0, []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                new, ok = fix(cell.value)
                if new != cell.value:
                    cell.value = new
                    changed += 1
                if not ok:
                    unknown.append(f"{ws.title}!{cell.coordinate}: {cell.value}")
wb.save(path)
print(f"Done. Fixed {changed} cells.")
if unknown:
    print("Could not auto-fix these - add them to MAPPING:")
    for u in unknown:
        print("  ", u)
input("Press Enter to close.")
