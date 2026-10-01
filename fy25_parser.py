import re
import pandas as pd
from collections import defaultdict
import pdfplumber

# -----------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------

PDF_FILE = "Activities Totals Report FY25.pdf"

# Patterns to detect activity start rows (ID - Name)
ACTIVITY_START = re.compile(r"^\d{5,6} - ")

# Patterns for season/category glued together
SEASON_CATEGORY = re.compile(
    r"(2024 Fall|2025 Spring|2025 Summer|2025 Fall|2024-25 Winter)"
    r"(Contract|Contracts|Tennis|Tennis & Pickleball|Tennis&Pickleball)"
)

# -----------------------------------------------------------
# STORAGE
# -----------------------------------------------------------

activities = []
current_block = []

# -----------------------------------------------------------
# FUNCTION TO FLUSH ACTIVITY BLOCKS
# -----------------------------------------------------------

def parse_block(block_lines):
    """Parse a multi-line activity block and return a dict."""

    full_text = " ".join(block_lines).replace("\n", " ")
    full_text = re.sub(r"\s+", " ", full_text).strip()

    # Detect season + category
    season_cat_match = SEASON_CATEGORY.search(full_text)
    if not season_cat_match:
        return None

    season = season_cat_match.group(1)
    category = season_cat_match.group(2)

    # Remove everything before season/category to isolate data row
    data_part = full_text.split(season + category, 1)[1].strip()

    # Extract numeric fields in order: Min, Max, Hours, Days, Res, NonRes
    nums = re.findall(r"[0-9]+(?:\.[0-9]+)?|N/A", data_part)

    if len(nums) < 6:
        return None

    # Assign fields (best effort)
    min_enroll = nums[0]
    max_enroll = nums[1]
    hours = nums[2]
    days = nums[3]
    res = nums[4]
    nonres = nums[5]
    total_enroll = int(res) + int(nonres) \
        if res.isdigit() and nonres.isdigit() else None

    # Activity name
    first_line = block_lines[0]
    activity_name = first_line.strip()

    return {
        "activity": activity_name,
        "season": season,
        "category": category,
        "min": min_enroll,
        "max": max_enroll,
        "hours": hours,
        "days": days,
        "res_enroll": res,
        "nonres_enroll": nonres,
        "total_enroll": total_enroll
    }

# -----------------------------------------------------------
# PARSE PDF
# -----------------------------------------------------------

with pdfplumber.open(PDF_FILE) as pdf:
    for page in pdf.pages:
        lines = page.extract_text().split("\n")
        for line in lines:
            if ACTIVITY_START.match(line):
                # flush previous
                if current_block:
                    parsed = parse_block(current_block)
                    if parsed:
                        activities.append(parsed)
                current_block = [line]
            else:
                # continue block
                if current_block:
                    current_block.append(line)

        # end-of-page flush
        if current_block:
            parsed = parse_block(current_block)
            if parsed:
                activities.append(parsed)
            current_block = []

# -----------------------------------------------------------
# BUILD DATAFRAME
# -----------------------------------------------------------

df = pd.DataFrame(activities)

# -----------------------------------------------------------
# SEASON & CATEGORY TOTALS
# -----------------------------------------------------------

season_totals = df.groupby("season").agg(
    classes=("activity", "count"),
    res_total=("res_enroll", lambda x: sum(int(i) for i in x if str(i).isdigit())),
    nonres_total=("nonres_enroll", lambda x: sum(int(i) for i in x if str(i).isdigit())),
    total_enroll=("total_enroll", "sum")
).reset_index()

category_totals = df.groupby("category").agg(
    classes=("activity", "count"),
    res_total=("res_enroll", lambda x: sum(int(i) for i in x if str(i).isdigit())),
    nonres_total=("nonres_enroll", lambda x: sum(int(i) for i in x if str(i).isdigit())),
    total_enroll=("total_enroll", "sum")
).reset_index()

grand_totals = {
    "classes": len(df),
    "res_total": season_totals["res_total"].sum(),
    "nonres_total": season_totals["nonres_total"].sum(),
    "total_enroll": season_totals["total_enroll"].sum()
}

# -----------------------------------------------------------
# PRINT PLAIN TEXT TOTALS
# -----------------------------------------------------------

print("==== SEASON TOTALS ====")
print(season_totals.to_string(index=False))

print("\n==== CATEGORY TOTALS ====")
print(category_totals.to_string(index=False))

print("\n==== GRAND TOTALS ====")
for k, v in grand_totals.items():
    print(f"{k}: {v}")

# -----------------------------------------------------------
# EXPORT TO EXCEL
# -----------------------------------------------------------

with pd.ExcelWriter("FY25_Extracted_Totals.xlsx") as writer:
    df.to_excel(writer, sheet_name="Parsed Activities", index=False)
    season_totals.to_excel(writer, sheet_name="Season Totals", index=False)
    category_totals.to_excel(writer, sheet_name="Category Totals", index=False)
    pd.DataFrame([grand_totals]).to_excel(writer, sheet_name="Grand Totals", index=False)

print("\nExcel file created: FY25_Extracted_Totals.xlsx")

with open("text_output.txt", "w") as f:
    f.write("==== SEASON TOTALS ====\n")
    f.write(season_totals.to_string(index=False))
    f.write("\n\n==== CATEGORY TOTALS ====\n")
    f.write(category_totals.to_string(index=False))
    f.write("\n\n==== GRAND TOTALS ====\n")
    for k, v in grand_totals.items():
        f.write(f"{k}: {v}\n")
