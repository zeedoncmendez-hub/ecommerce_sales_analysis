
from pathlib import Path

import nbformat as nbf
from nbclient import NotebookClient

ROOT = Path(__file__).resolve().parent
NB = ROOT / "notebooks" / "E-Commerce_Sales_Analysis.ipynb"

md = nbf.v4.new_markdown_cell
code = nbf.v4.new_code_cell

cells = [
    md("# E-Commerce Sales Analysis\n"
       "**Tools:** Python · pandas · NumPy · Matplotlib · Seaborn\n\n"
       "**Business questions**\n"
       "1. How much revenue does the business generate?\n"
       "2. Which products and categories perform best?\n"
       "3. How do sales change over time?\n"
       "4. Which regions and customers contribute most to sales?\n\n"
       "Workflow: **load → clean → explore → analyse → visualise → conclude**"),

    md("## 1. Load the dataset"),
    code("""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

sns.set_theme(style="whitegrid")

DATA = ROOT / "data" / "ecommerce_sales.csv"
df = pd.read_csv(DATA)

print("Dataset shape:", df.shape)
display(df.head())
"""),

    md("## 2. Check data quality"),
    code("""
print("Column names:")
print(df.columns.tolist())

print("\\nMissing values:")
display(df.isnull().sum().to_frame("Missing Values"))

print("Duplicate rows:", df.duplicated().sum())
print("\\nData types:")
display(df.dtypes.to_frame("Data Type"))
"""),

    md("## 3. Clean the data"),
    code("""
df.columns = (
    df.columns.str.strip()
      .str.lower()
      .str.replace(r"[^a-z0-9]+", "_", regex=True)
      .str.strip("_")
)

df = df.drop_duplicates().copy()

# Convert likely date columns
for col in df.columns:
    if "date" in col:
        df[col] = pd.to_datetime(df[col], errors="coerce")

# Convert likely numeric columns
for col in df.columns:
    if any(word in col for word in
           ["sales", "revenue", "price", "quantity", "profit", "amount"]):
        df[col] = pd.to_numeric(df[col], errors="coerce")

print("Rows after removing duplicates:", len(df))
display(df.head())
"""),

    md("## 4. Summary statistics"),
    code("""
display(df.describe(include="all").T)
"""),

    md("## 5. Calculate key business metrics"),
    code("""
# Update these column names if your dataset uses different names
sales_col = next(
    (c for c in ["sales", "revenue", "total_sales", "amount"]
     if c in df.columns),
    None
)

if sales_col:
    print("Total revenue:", round(df[sales_col].sum(), 2))
    print("Average sales:", round(df[sales_col].mean(), 2))
    print("Median sales:", round(df[sales_col].median(), 2))
else:
    print("Sales column not detected. Check df.columns.")
"""),

    md("## 6. Sales distribution"),
    code("""
if sales_col:
    sns.histplot(df[sales_col].dropna(), kde=True)
    plt.title("Sales Distribution")
    plt.xlabel(sales_col)
    plt.tight_layout()
    plt.show()
"""),

    md("## 7. Sales by category"),
    code("""
category_col = next(
    (c for c in ["category", "product_category", "category_name"]
     if c in df.columns),
    None
)

if sales_col and category_col:
    category_sales = (
        df.groupby(category_col)[sales_col]
          .sum()
          .sort_values(ascending=False)
    )

    display(category_sales.to_frame("Total Sales"))
    category_sales.plot(kind="bar")
    plt.title("Sales by Category")
    plt.ylabel("Total Sales")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()
else:
    print("Sales or category column not detected.")
"""),

    md("## 8. Monthly sales trend"),
    code("""
date_col = next(
    (c for c in ["order_date", "purchase_date", "date"]
     if c in df.columns),
    None
)

if sales_col and date_col:
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    monthly_sales = (
        df.dropna(subset=[date_col, sales_col])
          .assign(month=lambda x: x[date_col].dt.to_period("M").astype(str))
          .groupby("month")[sales_col]
          .sum()
    )

    display(monthly_sales.to_frame("Total Sales"))
    monthly_sales.plot(marker="o")
    plt.title("Monthly Sales Trend")
    plt.xlabel("Month")
    plt.ylabel("Total Sales")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()
else:
    print("Sales or date column not detected.")
"""),

    md("## 9. Top-selling products"),
    code("""
product_col = next(
    (c for c in ["product_name", "product", "item_name"]
     if c in df.columns),
    None
)

if sales_col and product_col:
    top_products = (
        df.groupby(product_col)[sales_col]
          .sum()
          .nlargest(10)
    )

    display(top_products.to_frame("Total Sales"))
    top_products.sort_values().plot(kind="barh")
    plt.title("Top 10 Products by Sales")
    plt.xlabel("Total Sales")
    plt.tight_layout()
    plt.show()
else:
    print("Sales or product column not detected.")
"""),

    md("## 10. Regional analysis"),
    code("""
region_col = next(
    (c for c in ["region", "city", "state", "country"]
     if c in df.columns),
    None
)

if sales_col and region_col:
    region_sales = (
        df.groupby(region_col)[sales_col]
          .sum()
          .sort_values(ascending=False)
    )

    display(region_sales.to_frame("Total Sales"))
    region_sales.plot(kind="bar")
    plt.title("Sales by Region")
    plt.ylabel("Total Sales")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()
else:
    print("Sales or region column not detected.")
"""),

    md("## 11. Correlation analysis"),
    code("""
numeric_df = df.select_dtypes(include="number")

if numeric_df.shape[1] >= 2:
    sns.heatmap(numeric_df.corr(), annot=True, cmap="coolwarm", fmt=".2f")
    plt.title("Correlation Between Numeric Columns")
    plt.tight_layout()
    plt.show()
else:
    print("Not enough numeric columns for correlation analysis.")
"""),

    md("## 12. Key findings"),
    code("""
print("E-COMMERCE SALES ANALYSIS")
print("=" * 30)

if sales_col:
    print(f"Total revenue: {df[sales_col].sum():,.2f}")

if category_col and sales_col:
    best_category = df.groupby(category_col)[sales_col].sum().idxmax()
    print("Top category:", best_category)

if product_col and sales_col:
    best_product = df.groupby(product_col)[sales_col].sum().idxmax()
    print("Top product:", best_product)

if date_col and sales_col:
    print("Date range:",
          df[date_col].min(), "to", df[date_col].max())

print("\\nAnalysis completed.")
"""),
]

nb = nbf.v4.new_notebook(
    cells=cells,
    metadata={
        "kernelspec": {
            "name": "python3",
            "display_name": "Python 3",
            "language": "python"
        }
    }
)

NB.parent.mkdir(parents=True, exist_ok=True)

# Save the notebook before execution
nbf.write(nb, NB)

# Execute cells and save their outputs
NotebookClient(
    nb,
    timeout=300,
    kernel_name="python3",
    resources={"metadata": {"path": str(ROOT)}},
    allow_errors=True
).execute()

nbf.write(nb, NB)
print("Created and executed:", NB)
