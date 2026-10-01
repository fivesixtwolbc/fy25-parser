import re
import pandas as pd
import pdfplumber

# -----------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------

PDF_FILE = "Activities Totals Report FY25.pdf"

# Detect the start of an activity (ID - Name)
ACTIVITY_START = re.compile(r"^\d{5,6} - ")

# Season + category always appear glued together (e.g., 2024-25 WinterTennis)
SEASON_CATEGORY = re.compile(
    r"(2024 Fall|2025 Spring|2025 Summer|2025 Fall|2024-25 Winter)([A-Za-z]+)"
)

# -----------------------------------------------------------
# STORAGE
# -----------------------------------------------------------

activities = []
current_block = []

# -----------------------------------------------------------
# FUNCTION TO PARSE MULTI-LINE ACTIVITY BLOCKS
# -----------------------------------------------------------

def parse_block(block_lines):
    """
    Parse a multi-line activity block and extract:
    - season
    - category
    - min, max, hours, days
    - res, nonres, total_enroll
    - activity name
    """

    # Combine lines and normalize spacing
    full_text = " ".join(block_lines)
    full_text = re.sub(r"\s+", " ", full_text).strip()

    # Locate season + category
    season_cat_match = SEASON_CATEGORY.search(full_text)
    if not season_cat_match:
        return None

    season = season_cat_match.group(1)
    category = season_cat_match.group(2)

    # Data portion AFTER season+category
    data_part = full_text.split(season + category, 1)[1].strip()

    # Extract numeric fields (ignores percentages and times)
    nums = re.findall(r"\b(N/A|[0-9]{1,3}(?:\.[0-9]+)?)\b", data_part)

    # Require enough numeric fields for Min, Max, Hours, Days, Res, NonRes
    if len(nums) < 6:
        return None

    min_enroll = nums[0]
    max_enroll = nums[1]
    hours = nums[2]
    days = nums[3]
    res = nums[4]
    nonres = nums[5]

    # Convert res/nonres to integers safely
    res_val = int(res) if res.isdigit() else 0
    nonres_val = int(nonres) if nonres.isdigit() else 0
    total_enroll = res_val + nonres_val

    # Construct activity name from first two lines
    activity_name = " ".join(block_lines[:2]).strip()

    return {
        "activity": activity_name,
        "season": season,
        "category": category,
        "min": min_enroll,
        "max": max_enroll,
        "hours": hours,
        "days": days,
        "res_enroll": res_val,
        "nonres_enroll": nonres_val,
        "total_enroll": total_enroll,
    }


# -----------------------------------------------------------
# PARSE PDF
# -----------------------------------------------------------

with pdfplumber.open(PDF_FILE) as pdf:
    for page in pdf.pages:
        page_text = page.extract_text()
        if not page_text:
            continue

        lines = page_text.split("\n")

        for line in lines:
            if ACTIVITY_START.match(line):
                # Flush previous entry
                if current_block:
                    parsed = parse_block(current_block)
                    if parsed:
                        activities.append(parsed)
                current_block = [line]
            else:
                if current_block:
                    current_block.append(line)

        # Flush at end of page
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
    res_total=("res_enroll", "sum"),
    nonres_total=("nonres_enroll", "sum"),
    total_enroll=("total_enroll", "sum"),
).reset_index()

category_totals = df.groupby("category").agg(
    classes=("activity", "count"),
    res_total=("res_enroll", "sum"),
    nonres_total=("nonres_enroll", "sum"),
    total_enroll=("total_enroll", "sum"),
).reset_index()

grand_totals = {
    "classes": len(df),
    "res_total": season_totals["res_total"].sum(),
    "nonres_total": season_totals["nonres_total"].sum(),
    "total_enroll": season_totals["total_enroll"].sum(),
}

# -----------------------------------------------------------
# OUTPUT TEXT TOTALS
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

# Also write plain text output for GitHub Actions
with open("text_output.txt", "w") as f:
    f.write("==== SEASON TOTALS ====\n")
    f.write(season_totals.to_string(index=False))
    f.write("\n\n==== CATEGORY TOTALS ====\n")
    f.write(category_totals.to_string(index=False))
    f.write("\n\n==== GRAND TOTALS ====\n")
    for k, v in grand_totals.items():
        f.write(f"{k}: {v}\n")
