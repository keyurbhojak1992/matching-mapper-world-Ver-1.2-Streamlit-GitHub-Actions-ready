import os
import re
import zipfile
import shutil
from tqdm import tqdm
import win32com.client as win32


input_zip_file = r"D:\suggested files datewise\15-07-2026 BID265\Sorted\sorted_output_fast_keep_order.zip"
output_zip_file = r"D:\suggested files datewise\15-07-2026 BID265\Sorted\Separated_Suggest_BID265_Stock.zip"


TEMP_ROOT = r"C:\ExcelSplitTemp"


def sanitize_filename(name):
    name = re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()
    name = name.rstrip(". ")
    return name[:120] if name else "Sheet"


def unique_path(folder, filename):
    base, ext = os.path.splitext(filename)
    path = os.path.join(folder, filename)
    counter = 1

    while os.path.exists(path):
        path = os.path.join(folder, f"{base}_{counter}{ext}")
        counter += 1

    return path


def split_workbook_with_excel(excel, workbook_path, output_folder):
    wb_name = sanitize_filename(os.path.splitext(os.path.basename(workbook_path))[0])
    wb_output_folder = os.path.abspath(os.path.join(output_folder, wb_name))
    os.makedirs(wb_output_folder, exist_ok=True)

    wb = excel.Workbooks.Open(
        Filename=os.path.abspath(workbook_path),
        UpdateLinks=0,
        ReadOnly=True,
        IgnoreReadOnlyRecommended=True
    )

    try:
        sheet_names = [ws.Name for ws in wb.Worksheets]

        for sheet_name in tqdm(sheet_names, desc=wb_name, leave=False):
            ws = wb.Worksheets(sheet_name)

            ws.Copy()
            new_wb = excel.ActiveWorkbook

            safe_name = sanitize_filename(sheet_name)
            output_file = unique_path(wb_output_folder, f"{safe_name}.xlsx")
            output_file = os.path.abspath(output_file)

            try:
                new_wb.SaveAs(
                    Filename=output_file,
                    FileFormat=51,
                    ConflictResolution=2
                )
            finally:
                new_wb.Close(SaveChanges=False)

    finally:
        wb.Close(SaveChanges=False)

    return wb_output_folder


def process_zip():
    extract_folder = os.path.join(TEMP_ROOT, "extract")
    separated_folder = os.path.join(TEMP_ROOT, "separated")
    final_folder = os.path.join(TEMP_ROOT, "final")

    shutil.rmtree(TEMP_ROOT, ignore_errors=True)

    os.makedirs(extract_folder, exist_ok=True)
    os.makedirs(separated_folder, exist_ok=True)
    os.makedirs(final_folder, exist_ok=True)

    print("Step 1: Extracting master zip...")
    with zipfile.ZipFile(input_zip_file, "r") as z:
        z.extractall(extract_folder)

    print("Step 2: Finding Excel files...")
    excel_files = []

    for root, dirs, files in os.walk(extract_folder):
        for file in files:
            if file.startswith("~$"):
                continue
            if file.lower().endswith((".xlsx", ".xlsm")):
                excel_files.append(os.path.abspath(os.path.join(root, file)))

    print(f"Found {len(excel_files)} workbook(s).")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False

    try:
        print("Step 3: Splitting sheets using Excel...")
        for workbook_path in excel_files:
            folder = split_workbook_with_excel(excel, workbook_path, separated_folder)

            zip_name = os.path.basename(folder)
            shutil.make_archive(
                os.path.join(final_folder, zip_name),
                "zip",
                folder
            )

    finally:
        excel.Quit()

    print("Step 4: Creating final zip...")

    output_base = output_zip_file.replace(".zip", "")
    shutil.make_archive(output_base, "zip", final_folder)

    shutil.rmtree(TEMP_ROOT, ignore_errors=True)

    print(f"DONE: {output_zip_file}")


if os.path.exists(input_zip_file):
    process_zip()
else:
    print("Input zip file not found.")