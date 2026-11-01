import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Read original scraped data
df = pd.read_csv("stage_pfe_casablanca.csv")

# Select only the relevant, informative columns
relevant_cols = [
    "title",
    "company",
    "location",
    "date_posted",
    "site",
    "job_url",
    "job_type",
    "is_remote",
]

# Keep only columns that exist
existing_cols = [c for c in relevant_cols if c in df.columns]
clean_df = df[existing_cols].copy()

# Rename columns to user-friendly titles
rename_dict = {
    "title": "Job Title",
    "company": "Company",
    "location": "Location",
    "date_posted": "Date Posted",
    "site": "Job Board",
    "job_url": "Job Link",
    "job_type": "Job Type",
    "is_remote": "Remote",
}
clean_df.rename(columns=rename_dict, inplace=True)

# Fill NaNs with clean text
clean_df["Company"] = clean_df["Company"].fillna("Not Specified")
clean_df["Job Type"] = clean_df["Job Type"].fillna("Internship / PFE")
clean_df["Remote"] = clean_df["Remote"].apply(lambda x: "Yes" if x is True else ("No" if x is False else "Not Specified"))

# Sort by date posted (newest first)
if "Date Posted" in clean_df.columns:
    clean_df.sort_values(by="Date Posted", ascending=False, inplace=True)

# 1. Save cleaned standard CSV
clean_df.to_csv("stage_pfe_casablanca_clean.csv", index=False, encoding="utf-8-sig")

# 2. Save European/French Excel-compatible CSV (semicolon separated)
clean_df.to_csv("stage_pfe_casablanca_french_excel.csv", sep=";", index=False, encoding="utf-8-sig")

# 3. Create a polished, professionally formatted Excel workbook (.xlsx)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Stage PFE Casablanca"

# Styling definitions
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
cell_font = Font(name="Calibri", size=10)
link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

# Write headers
headers = list(clean_df.columns)
for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Write rows
for row_idx, row_data in enumerate(clean_df.itertuples(index=False), 2):
    for col_idx, val in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        header_name = headers[col_idx - 1]
        
        if header_name == "Job Link" and pd.notna(val) and str(val).startswith("http"):
            cell.value = "Click to View Listing"
            cell.hyperlink = str(val)
            cell.font = link_font
        else:
            cell.value = str(val) if pd.notna(val) else ""
            cell.font = cell_font
            
        cell.border = thin_border
        
        if header_name in ["Date Posted", "Job Board", "Remote", "Job Type"]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

# Auto-adjust column widths
for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

# Freeze top row and add filter
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

wb.save("stage_pfe_casablanca_clean.xlsx")
print("Successfully generated clean CSV and clean Excel workbook!")
