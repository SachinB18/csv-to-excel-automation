import pandas as pd

# -----------------------------
# STEP 1: Read CSV with encoding
# -----------------------------
df = pd.read_csv("raw_sales_data.csv", encoding="latin1")

print("Data loaded successfully")
print("Number of rows:", len(df))

# -----------------------------
# STEP 2: Clean column names
# -----------------------------
df.columns = df.columns.str.strip()

# -----------------------------
# STEP 3: Clean text columns
# -----------------------------
text_columns = df.select_dtypes(include="object").columns
for col in text_columns:
    df[col] = df[col].astype(str).str.strip()

# Replace 'nan' strings back to empty cells
df.replace("nan", "", inplace=True)

print("Text fields cleaned")

# -----------------------------
# STEP 4: Fix dates safely
# -----------------------------
if "ORDERDATE" in df.columns:
    df["ORDERDATE"] = pd.to_datetime(df["ORDERDATE"], errors="coerce")
    df["ORDERDATE"] = df["ORDERDATE"].dt.strftime("%Y-%m-%d")


print("Dates cleaned")

# -----------------------------
# STEP 5: Fix numeric columns
# -----------------------------
numeric_columns = df.select_dtypes(include="number").columns
df[numeric_columns] = df[numeric_columns].fillna(0)

print("Numeric values cleaned")

# -----------------------------
# STEP 6: Business summary
# -----------------------------
summary = None
if "SALES" in df.columns:
    summary = pd.DataFrame({
        "Metric": ["Total Sales", "Average Sales"],
        "Value": [
            round(df["SALES"].sum(), 2),
            round(df["SALES"].mean(), 2)
        ]
    })

print("Business summary generated")

# -----------------------------
# STEP 7: Export to Excel
# -----------------------------
with pd.ExcelWriter("final_report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    if summary is not None:
        summary.to_excel(writer, sheet_name="Summary", index=False)

print("Final Excel report generated: final_report.xlsx")
