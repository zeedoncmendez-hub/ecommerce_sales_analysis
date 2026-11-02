"""
E-Commerce Sales Analysis Excel Report
--------------------------------------
Creates a formatted Excel workbook with:
- Clean_Data source table
- Formula-driven dashboard and summary sheets
- Descriptive statistics with live Excel formulas
- Python/SciPy hypothesis test results (static, clearly labelled)

Requirements:
    python -m pip install pandas numpy scipy openpyxl

Run from your project folder:
    python src/create_excel_report.py

Expected input:
    Data/ecommerce_sales_raw.csv
Output:
    outputs/ecommerce_sales_analysis_report.xlsx

If your raw CSV is stored elsewhere, change RAW_DATA_PATH below.
"""

from pathlib import Path
import pandas as pd
import numpy as np
from scipy import stats
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------
# 1. Paths
# ---------------------------------------------------------------------
PROJECT_DIR = Path(__file__).resolve().parents[1]
RAW_DATA_PATH = PROJECT_DIR / "Data" / "raw" / "ecommerce_sales_raw.csv"
OUTPUT_DIR = PROJECT_DIR / "outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_PATH = OUTPUT_DIR / "ecommerce_sales_analysis_report.xlsx"

if not RAW_DATA_PATH.exists():
    raise FileNotFoundError(
        f"Could not find raw dataset: {RAW_DATA_PATH}\n"
        "Place ecommerce_sales_raw.csv in your project's Data folder, "
        "or update RAW_DATA_PATH in this script."
    )


# ---------------------------------------------------------------------
# 2. Read and clean data
# ---------------------------------------------------------------------
raw = pd.read_csv(RAW_DATA_PATH)
df = raw.drop_duplicates().copy()

text_cols = [
    "gender", "city", "region", "category",
    "sub_category", "payment_method"
]

for col in text_cols:
    if col in df.columns:
        df[col] = (
            df[col].astype("string").str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )

for col in ["gender", "city", "region", "category", "sub_category"]:
    if col in df.columns:
        df[col] = df[col].str.title()

if "payment_method" in df.columns:
    df["payment_method"] = df["payment_method"].str.title()
    df["payment_method"] = df["payment_method"].replace({
        "Upi": "UPI",
        "Cod": "COD",
    })

df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")

numeric_cols = [
    "customer_age", "unit_price", "quantity",
    "discount_pct", "delivery_days", "rating"
]
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Replace invalid values with missing values
df.loc[~df["customer_age"].between(13, 100), "customer_age"] = np.nan
df.loc[df["unit_price"] <= 0, "unit_price"] = np.nan
df.loc[df["quantity"] <= 0, "quantity"] = np.nan
df.loc[~df["discount_pct"].between(0, 100), "discount_pct"] = np.nan
df.loc[df["delivery_days"] < 1, "delivery_days"] = np.nan
df.loc[~df["rating"].between(1, 5), "rating"] = np.nan

# Fill numeric missing values with the median
for col in numeric_cols:
    if df[col].notna().any():
        df[col] = df[col].fillna(df[col].median())

# Fill categorical missing values with the mode
for col in text_cols:
    if col in df.columns and df[col].notna().any():
        modes = df[col].mode(dropna=True)
        if not modes.empty:
            df[col] = df[col].fillna(modes.iloc[0])

# Required fields
df = df.dropna(subset=["order_id", "order_date", "customer_id"]).copy()

# Create analysis columns
df["sales_amount"] = (
    df["unit_price"] * df["quantity"]
    * (1 - df["discount_pct"] / 100)
)
df["month_start"] = df["order_date"].dt.to_period("M").dt.to_timestamp()

# Normalize return field
df["returned"] = (
    df["returned"].astype(str).str.strip().str.lower()
    .map({
        "true": True, "false": False,
        "1": True, "0": False,
        "yes": True, "no": False
    })
)
# Missing return flags are left as False for this report's return-rate measure.
df["returned"] = df["returned"].fillna(False).astype(bool)

# Make values Excel-friendly
df = df.replace({np.nan: None})


# ---------------------------------------------------------------------
# 3. Workbook and theme
# ---------------------------------------------------------------------
wb = Workbook()
ws_dash = wb.active
ws_dash.title = "Dashboard"
ws_clean = wb.create_sheet("Clean_Data")
ws_cat = wb.create_sheet("Category_Summary")
ws_month = wb.create_sheet("Monthly_Summary")
ws_region = wb.create_sheet("Region_Summary")
ws_stats = wb.create_sheet("Descriptive_Stats")
ws_tests = wb.create_sheet("Hypothesis_Tests")
ws_readme = wb.create_sheet("Read_Me")

NAVY = "17365D"
BLUE = "4472C4"
PALE_BLUE = "EAF2F8"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
GRAY = "666666"
ORANGE = "FCE4D6"
GREEN = "E2F0D9"
thin_gray = Side(style="thin", color="D9E1F2")
medium_blue = Side(style="medium", color=BLUE)


def title_block(ws, title, subtitle, end_col=6):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    c = ws.cell(1, 1, title)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    c = ws.cell(2, 1, subtitle)
    c.font = Font(name="Aptos", size=10, italic=True, color=GRAY)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 30


def style_header(ws, row, start_col, end_col):
    for col_num in range(start_col, end_col + 1):
        cell = ws.cell(row, col_num)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = Border(bottom=medium_blue)
    ws.row_dimensions[row].height = 28


def setup_sheet(ws, freeze="A4"):
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------
# 4. Clean_Data worksheet
# ---------------------------------------------------------------------
headers = list(df.columns)
for col_num, header in enumerate(headers, start=1):
    ws_clean.cell(1, col_num, header)

for row_num, row in enumerate(df.itertuples(index=False, name=None), start=2):
    for col_num, value in enumerate(row, start=1):
        if value is None or pd.isna(value):
            value = None
        elif isinstance(value, pd.Timestamp):
            value = value.to_pydatetime()
        elif isinstance(value, (np.bool_, bool)):
            value = bool(value)
        elif isinstance(value, (np.integer,)):
            value = int(value)
        elif isinstance(value, (np.floating, float)):
            value = float(value)
        ws_clean.cell(row_num, col_num, value)

style_header(ws_clean, 1, 1, len(headers))
setup_sheet(ws_clean, "A2")

last_data_row = len(df) + 1
last_col_letter = get_column_letter(len(headers))
ws_clean.auto_filter.ref = f"A1:{last_col_letter}{last_data_row}"

column_number = {name: i + 1 for i, name in enumerate(headers)}
for col_num, header in enumerate(headers, start=1):
    ws_clean.column_dimensions[get_column_letter(col_num)].width = min(
        max(len(header) + 3, 13), 22
    )

for r in range(2, last_data_row + 1):
    if "order_date" in column_number:
        ws_clean.cell(r, column_number["order_date"]).number_format = "dd-mmm-yyyy"
    if "month_start" in column_number:
        ws_clean.cell(r, column_number["month_start"]).number_format = "mmm yyyy"
    for name in ["unit_price", "sales_amount"]:
        if name in column_number:
            ws_clean.cell(r, column_number[name]).number_format = '₹#,##0.00'
    if "discount_pct" in column_number:
        ws_clean.cell(r, column_number["discount_pct"]).number_format = "0.0"
    if "rating" in column_number:
        ws_clean.cell(r, column_number["rating"]).number_format = "0.0"

table = Table(
    displayName="CleanData",
    ref=f"A1:{last_col_letter}{last_data_row}"
)
table.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2",
    showRowStripes=True,
    showColumnStripes=False
)
ws_clean.add_table(table)

# Excel formula ranges refer to the current populated rows of Clean_Data.
col = {name: get_column_letter(i + 1) for i, name in enumerate(headers)}
def data_range(field):
    return f"Clean_Data!${col[field]}$2:${col[field]}${last_data_row}"

sales_rng = data_range("sales_amount")
order_rng = data_range("order_id")
customer_rng = data_range("customer_id")
quantity_rng = data_range("quantity")
category_rng = data_range("category")
region_rng = data_range("region")
payment_rng = data_range("payment_method")
rating_rng = data_range("rating")
delivery_rng = data_range("delivery_days")
returned_rng = data_range("returned")
month_rng = data_range("month_start")


# ---------------------------------------------------------------------
# 5. Dashboard — live Excel formulas
# ---------------------------------------------------------------------
title_block(
    ws_dash,
    "E-Commerce Sales Analysis Dashboard",
    "Formula-driven report. Summary formulas reference Clean_Data and recalculate in Excel.",
    8
)
setup_sheet(ws_dash, "A4")

for j, value in enumerate(["KPI", "Live Excel Formula", "Interpretation"], 1):
    ws_dash.cell(4, j, value)
style_header(ws_dash, 4, 1, 3)

kpis = [
    ("Total Revenue", f"=SUM({sales_rng})", "Net sales after discounts"),
    ("Unique Orders", f"=COUNTA(UNIQUE({order_rng}))", "Distinct order IDs (Excel 2021 / Microsoft 365)"),
    ("Unique Customers", f"=COUNTA(UNIQUE({customer_rng}))", "Distinct customer IDs (Excel 2021 / Microsoft 365)"),
    ("Units Sold", f"=SUM({quantity_rng})", "Total units sold"),
    ("Average Order Value", "=IFERROR(B5/B6,0)", "Revenue divided by unique orders"),
    ("Average Rating", f"=AVERAGE({rating_rng})", "Mean rating"),
    ("Return Rate", f'=IFERROR(COUNTIFS({returned_rng},TRUE)/COUNTA({returned_rng}),0)', "Share of rows marked returned"),
    ("Average Delivery Days", f"=AVERAGE({delivery_rng})", "Mean delivery time"),
]
for r, (label, formula, note) in enumerate(kpis, start=5):
    ws_dash.cell(r, 1, label)
    ws_dash.cell(r, 2, formula)
    ws_dash.cell(r, 3, note)
    ws_dash.cell(r, 1).font = Font(bold=True, color=NAVY)
    ws_dash.cell(r, 2).fill = PatternFill("solid", fgColor=PALE_BLUE)
    ws_dash.cell(r, 3).font = Font(color=GRAY, italic=True)
    ws_dash.cell(r, 2).number_format = (
        '₹#,##0.00' if label in ("Total Revenue", "Average Order Value")
        else "0.0%" if label == "Return Rate"
        else "0.00" if label in ("Average Rating", "Average Delivery Days")
        else "#,##0"
    )

ws_dash["A15"] = "Formula compatibility"
ws_dash["A15"].font = Font(bold=True, color=NAVY)
ws_dash.merge_cells("A16:H17")
ws_dash["A16"] = (
    "Unique order/customer KPIs use UNIQUE(), available in Excel 2021 and Microsoft 365. "
    "For older versions, use a PivotTable or a compatible distinct-count formula."
)
ws_dash["A16"].alignment = Alignment(wrap_text=True, vertical="top")
ws_dash["A16"].fill = PatternFill("solid", fgColor=LIGHT_GRAY)
ws_dash.column_dimensions["A"].width = 26
ws_dash.column_dimensions["B"].width = 24
ws_dash.column_dimensions["C"].width = 54


# ---------------------------------------------------------------------
# 6. Category summary
# ---------------------------------------------------------------------
title_block(
    ws_cat, "Category Performance",
    "Live formulas summarize revenue, volume, ratings and returns by category.", 7
)
setup_sheet(ws_cat)
cat_headers = [
    "Category", "Revenue (₹)", "Units Sold", "Unique Orders",
    "Average Rating", "Returned Rows", "Return Rate"
]
for j, h in enumerate(cat_headers, 1):
    ws_cat.cell(4, j, h)
style_header(ws_cat, 4, 1, len(cat_headers))

categories = sorted(df["category"].dropna().astype(str).unique())
for r, category in enumerate(categories, 5):
    ws_cat.cell(r, 1, category)
    ws_cat.cell(r, 2, f'=SUMIFS({sales_rng},{category_rng},$A{r})')
    ws_cat.cell(r, 3, f'=SUMIFS({quantity_rng},{category_rng},$A{r})')
    # Distinct orders within category; formula-based and recalculates in Excel.
    ws_cat.cell(
        r, 4,
        f'=SUMPRODUCT(({category_rng}=$A{r})/COUNTIFS({order_rng},{order_rng},{category_rng},$A{r}))'
    )
    ws_cat.cell(r, 5, f'=AVERAGEIFS({rating_rng},{category_rng},$A{r})')
    ws_cat.cell(r, 6, f'=COUNTIFS({category_rng},$A{r},{returned_rng},TRUE)')
    ws_cat.cell(r, 7, f'=IFERROR(F{r}/COUNTIF({category_rng},$A{r}),0)')
    ws_cat.cell(r, 2).number_format = '₹#,##0.00'
    ws_cat.cell(r, 5).number_format = "0.00"
    ws_cat.cell(r, 7).number_format = "0.0%"

if categories:
    ws_cat.conditional_formatting.add(
        f"B5:B{4 + len(categories)}",
        ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B"
        )
    )

for letter, width in {
    "A": 23, "B": 20, "C": 15, "D": 17, "E": 18, "F": 16, "G": 15
}.items():
    ws_cat.column_dimensions[letter].width = width


# ---------------------------------------------------------------------
# 7. Monthly summary and line chart
# ---------------------------------------------------------------------
title_block(
    ws_month, "Monthly Sales Trend",
    "Monthly revenue, row counts and average delivery time from Clean_Data.", 5
)
setup_sheet(ws_month)
month_headers = [
    "Month", "Revenue (₹)", "Order Rows",
    "Average Row Revenue", "Average Delivery Days"
]
for j, h in enumerate(month_headers, 1):
    ws_month.cell(4, j, h)
style_header(ws_month, 4, 1, len(month_headers))

months = sorted(pd.to_datetime(df["month_start"].dropna().unique()))
for r, month in enumerate(months, 5):
    month = pd.Timestamp(month).to_pydatetime()
    ws_month.cell(r, 1, month)
    ws_month.cell(r, 1).number_format = "mmm yyyy"
    ws_month.cell(r, 2, f"=SUMIFS({sales_rng},{month_rng},$A{r})")
    ws_month.cell(r, 3, f"=COUNTIFS({month_rng},$A{r})")
    ws_month.cell(r, 4, f"=IFERROR(B{r}/C{r},0)")
    ws_month.cell(r, 5, f"=AVERAGEIFS({delivery_rng},{month_rng},$A{r})")
    ws_month.cell(r, 2).number_format = '₹#,##0.00'
    ws_month.cell(r, 4).number_format = '₹#,##0.00'
    ws_month.cell(r, 5).number_format = "0.0"

if months:
    chart = LineChart()
    chart.title = "Monthly Revenue"
    chart.style = 13
    chart.y_axis.title = "Revenue (₹)"
    chart.x_axis.title = "Month"
    chart.height = 7
    chart.width = 14
    chart_data = Reference(ws_month, min_col=2, min_row=4, max_row=4 + len(months))
    chart_categories = Reference(ws_month, min_col=1, min_row=5, max_row=4 + len(months))
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(chart_categories)
    chart.legend = None
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    ws_month.add_chart(chart, "G4")

for letter, width in {
    "A": 16, "B": 20, "C": 15, "D": 23, "E": 23
}.items():
    ws_month.column_dimensions[letter].width = width


# ---------------------------------------------------------------------
# 8. Regional summary
# ---------------------------------------------------------------------
title_block(
    ws_region, "Regional Performance",
    "Live formulas compare revenue, units, row counts and delivery times by region.", 5
)
setup_sheet(ws_region)
region_headers = [
    "Region", "Revenue (₹)", "Units Sold", "Order Rows", "Average Delivery Days"
]
for j, h in enumerate(region_headers, 1):
    ws_region.cell(4, j, h)
style_header(ws_region, 4, 1, len(region_headers))

regions = sorted(df["region"].dropna().astype(str).unique())
for r, region in enumerate(regions, 5):
    ws_region.cell(r, 1, region)
    ws_region.cell(r, 2, f"=SUMIFS({sales_rng},{region_rng},$A{r})")
    ws_region.cell(r, 3, f"=SUMIFS({quantity_rng},{region_rng},$A{r})")
    ws_region.cell(r, 4, f"=COUNTIFS({region_rng},$A{r})")
    ws_region.cell(r, 5, f"=AVERAGEIFS({delivery_rng},{region_rng},$A{r})")
    ws_region.cell(r, 2).number_format = '₹#,##0.00'
    ws_region.cell(r, 5).number_format = "0.0"

for letter, width in {
    "A": 22, "B": 20, "C": 15, "D": 15, "E": 25
}.items():
    ws_region.column_dimensions[letter].width = width


# ---------------------------------------------------------------------
# 9. Descriptive statistics — live Excel formulas
# ---------------------------------------------------------------------
title_block(
    ws_stats, "Descriptive Statistics",
    "Excel formulas recalculate descriptive statistics from Clean_Data.", 3
)
setup_sheet(ws_stats)
for j, h in enumerate(["Metric", "Live Excel Formula", "Description"], 1):
    ws_stats.cell(4, j, h)
style_header(ws_stats, 4, 1, 3)

descriptive = [
    ("Mean row revenue", f"=AVERAGE({sales_rng})", "Average sales amount per row"),
    ("Median row revenue", f"=MEDIAN({sales_rng})", "Middle row revenue value"),
    ("Sample standard deviation", f"=STDEV({sales_rng})", "Spread of row revenue values"),
    ("Minimum row revenue", f"=MIN({sales_rng})", "Smallest row revenue"),
    ("Maximum row revenue", f"=MAX({sales_rng})", "Largest row revenue"),
    ("Mean quantity", f"=AVERAGE({quantity_rng})", "Average units per row"),
    ("Median quantity", f"=MEDIAN({quantity_rng})", "Median units per row"),
    ("Mean rating", f"=AVERAGE({rating_rng})", "Average rating"),
    ("Mean delivery days", f"=AVERAGE({delivery_rng})", "Average delivery time"),
]
for r, (metric, formula, description) in enumerate(descriptive, 5):
    ws_stats.cell(r, 1, metric)
    ws_stats.cell(r, 2, formula)
    ws_stats.cell(r, 3, description)
    ws_stats.cell(r, 2).fill = PatternFill("solid", fgColor=PALE_BLUE)
    ws_stats.cell(r, 2).number_format = (
        '₹#,##0.00' if "revenue" in metric.lower()
        else "0.00"
    )

ws_stats.column_dimensions["A"].width = 31
ws_stats.column_dimensions["B"].width = 27
ws_stats.column_dimensions["C"].width = 43


# ---------------------------------------------------------------------
# 10. Hypothesis test — Python/SciPy static results
# ---------------------------------------------------------------------
title_block(
    ws_tests,
    "Hypothesis Tests — Python / SciPy Results",
    "Values below are calculated in Python using scipy.stats. They are static, not live Excel formulas.",
    6
)
setup_sheet(ws_tests)
test_headers = [
    "Test", "Null Hypothesis", "Comparison",
    "Test Statistic", "p-value", "Decision (alpha = 0.05)"
]
for j, h in enumerate(test_headers, 1):
    ws_tests.cell(4, j, h)
style_header(ws_tests, 4, 1, len(test_headers))

payment_counts = (
    df.groupby("payment_method")["order_id"]
    .nunique().sort_values(ascending=False)
)
eligible_methods = payment_counts[payment_counts >= 2].index.tolist()

if len(eligible_methods) >= 2:
    method_a, method_b = eligible_methods[:2]
    group_a = df.loc[
        df["payment_method"] == method_a, "sales_amount"
    ].dropna().astype(float).to_numpy()
    group_b = df.loc[
        df["payment_method"] == method_b, "sales_amount"
    ].dropna().astype(float).to_numpy()

    if len(group_a) >= 2 and len(group_b) >= 2:
        t_stat, p_value = stats.ttest_ind(
            group_a, group_b, equal_var=False, nan_policy="omit"
        )
        decision = "Reject H0" if p_value < 0.05 else "Fail to reject H0"
        test_values = [
            "Welch's independent-samples t-test",
            "Mean row revenue is equal between payment groups",
            f"{method_a} vs {method_b}",
            float(t_stat),
            float(p_value),
            decision
        ]
        for j, value in enumerate(test_values, 1):
            ws_tests.cell(5, j, value)
        ws_tests["D5"].number_format = "0.0000"
        ws_tests["E5"].number_format = "0.000000"
    else:
        ws_tests["A5"] = "Not run"
        ws_tests["B5"] = "Each group needs at least two observations."
else:
    ws_tests["A5"] = "Not run"
    ws_tests["B5"] = "At least two eligible payment groups are required."

ws_tests.merge_cells("A8:F8")
ws_tests["A8"] = (
    "Method: two-sided Welch t-test using scipy.stats.ttest_ind(equal_var=False)."
)
ws_tests["A8"].font = Font(italic=True, color=GRAY)
ws_tests.merge_cells("A9:F10")
ws_tests["A9"] = (
    "Interpretation: p < 0.05 indicates evidence of a difference in mean row revenue "
    "between the groups. A non-significant result does not prove that the means are equal. "
    "The dataset is synthetic and used for demonstration."
)
ws_tests["A9"].alignment = Alignment(wrap_text=True, vertical="top")
ws_tests["A9"].fill = PatternFill("solid", fgColor=ORANGE)

for letter, width in {
    "A": 37, "B": 46, "C": 23, "D": 17, "E": 17, "F": 25
}.items():
    ws_tests.column_dimensions[letter].width = width


# ---------------------------------------------------------------------
# 11. Read_Me guide
# ---------------------------------------------------------------------
title_block(
    ws_readme, "Workbook Guide & Assumptions",
    "Review these notes before presenting the report.", 2
)
setup_sheet(ws_readme)
ws_readme["A4"] = "Topic"
ws_readme["B4"] = "Details"
style_header(ws_readme, 4, 1, 2)

guide = [
    ("Purpose", "E-commerce sales report with formula-driven Excel summaries and a Python-calculated hypothesis test."),
    ("Clean_Data", "Cleaned, standardized records used by summary formulas. Edit existing values here to update formula-based summaries."),
    ("Live formulas", "Dashboard, category, monthly, regional and descriptive-statistics sheets contain Excel formulas linked to Clean_Data."),
    ("Excel version", "Dashboard unique counts use UNIQUE(), available in Microsoft 365 / Excel 2021+."),
    ("Hypothesis tests", "The test statistic and p-value are static Python/SciPy results. Re-run this script after changing source data."),
    ("Revenue definition", "sales_amount = unit_price × quantity × (1 − discount_pct / 100)."),
    ("Return rate", "Return rate is the proportion of cleaned rows marked returned, not necessarily a unique-order return rate."),
    ("Data source", "Synthetic India-oriented e-commerce sample data, not actual Amazon or Flipkart records."),
    ("Recalculation", "Excel is set to automatic recalculation on open. If needed, use Formulas → Calculation Options → Automatic."),
    ("Adding rows", "Summary formula ranges are based on the data row count at workbook creation. If you add rows below the current range, rerun this script to extend formulas."),
]
for r, (topic, details) in enumerate(guide, 5):
    ws_readme.cell(r, 1, topic).font = Font(bold=True, color=NAVY)
    ws_readme.cell(r, 2, details).alignment = Alignment(wrap_text=True, vertical="top")
    ws_readme.row_dimensions[r].height = 34 if len(details) > 100 else 26

ws_readme.column_dimensions["A"].width = 23
ws_readme.column_dimensions["B"].width = 105


# ---------------------------------------------------------------------
# 12. Global formatting and formula recalculation settings
# ---------------------------------------------------------------------
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25

try:
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
except Exception:
    pass

wb.save(OUTPUT_PATH)

# Verify that the workbook was saved and contains formulas.
check = load_workbook(OUTPUT_PATH, data_only=False)
formula_count = sum(
    1
    for sheet_name in [
        "Dashboard", "Category_Summary", "Monthly_Summary",
        "Region_Summary", "Descriptive_Stats"
    ]
    for row in check[sheet_name].iter_rows()
    for cell in row
    if isinstance(cell.value, str) and cell.value.startswith("=")
)

print(f"Workbook created: {OUTPUT_PATH}")
print(f"Raw rows: {len(raw):,}")
print(f"Clean rows: {len(df):,}")
print(f"Sheets: {', '.join(check.sheetnames)}")
print(f"Live Excel formulas: {formula_count}")
print("Hypothesis-test results are static Python/SciPy outputs.")
