import pandas as pd

# ==============================
# STEP 1: LOAD CLEAN DATA
# ==============================
df = pd.read_csv("sales final.csv")

# ==============================
# STEP 2: KPI CALCULATIONS
# ==============================
total_sales = df["sales"].sum()
avg_sales = df["sales"].mean()

top_country = df.groupby("country")["sales"].sum().idxmax()
top_product = df.groupby("productline")["sales"].sum().idxmax()

# ==============================
# STEP 3: KPI SUMMARY TABLE
# ==============================
kpi_df = pd.DataFrame({
    "KPI": [
        "TOTAL SALES",
        "AVERAGE SALES PER TRANSACTION",
        "TOP COUNTRY",
        "TOP PRODUCT LINE"
    ],
    "VALUE": [
        round(total_sales, 2),
        round(avg_sales, 2),
        top_country,
        top_product
    ]
})

# ==============================
# STEP 4: AGGREGATIONS
# ==============================
sales_by_year = df.groupby("year", as_index=False)["sales"].sum()
sales_by_country = df.groupby("country", as_index=False)["sales"].sum()
sales_by_product = df.groupby("productline", as_index=False)["sales"].sum()
sales_by_dealsize = df.groupby("dealsize", as_index=False)["sales"].sum()

# ==============================
# STEP 5: EXPORT TO EXCEL
# ==============================
with pd.ExcelWriter("automated_sales_report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="CLEAN_DATA", index=False)
    kpi_df.to_excel(writer, sheet_name="KPI_SUMMARY", index=False)
    sales_by_year.to_excel(writer, sheet_name="SALES_BY_YEAR", index=False)
    sales_by_country.to_excel(writer, sheet_name="SALES_BY_COUNTRY", index=False)
    sales_by_product.to_excel(writer, sheet_name="SALES_BY_PRODUCT", index=False)
    sales_by_dealsize.to_excel(writer, sheet_name="SALES_BY_DEALSIZE", index=False)

print("Automation complete. File saved as automated_sales_report.xlsx")
