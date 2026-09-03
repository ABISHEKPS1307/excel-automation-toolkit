
import os
import pandas as pd
from openpyxl import load_workbook

INPUT_FOLDER = "input"
OUTPUT_FOLDER = "output"

os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

files = [f for f in os.listdir(INPUT_FOLDER) if f.endswith(".xlsx")]

if not files:
    print("No Excel file found inside the input folder.")
    exit()

file_name = files[0]
input_path = os.path.join(INPUT_FOLDER, file_name)

output_name = file_name.replace(".xlsx", "_cleaned.xlsx")
output_path = os.path.join(OUTPUT_FOLDER, output_name)

print(f"Found: {file_name}")
print("Cleaning...")

df = pd.read_excel(input_path)

original_rows = len(df)

df = df.drop_duplicates()
df = df.dropna(how="all")

if "Name" in df.columns:
    df = df.dropna(subset=["Name"])

removed_rows = original_rows - len(df)

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)

    summary = pd.DataFrame({
        "Metric": [
            "Original Rows",
            "Final Rows",
            "Removed Rows",
            "Total Sales"
        ],
        "Value": [
            original_rows,
            len(df),
            removed_rows,
            df["Amount"].sum() if "Amount" in df.columns else "N/A"
        ]
    })

    summary.to_excel(writer, sheet_name="Summary", index=False)

wb = load_workbook(output_path)

for sheet in wb.worksheets:
    for column in sheet.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
        sheet.column_dimensions[column[0].column_letter].width = max_length + 3

wb.save(output_path)

print("Done!")
print(f"Saved to: {output_path}")