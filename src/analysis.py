
import pandas as pd
import numpy as np
import statistics
from scipy import stats
import matplotlib.pyplot as plt

# 1. Load cleaned data
df = pd.read_csv(
    r"Data/Processed/ecommerce_sales_cleaned.csv"
)

df["order_date"] = pd.to_datetime(df["order_date"])

print("\n--- DATA OVERVIEW ---")
print("Rows and columns:", df.shape)
print(df.head())
print(df.isnull().sum())

# 2. Create sales and revenue metrics
df["sales_amount"] = (
    df["unit_price"]
    * df["quantity"]
    * (1 - df["discount_pct"] / 100)
)

# 3. KPIs
total_orders = df["order_id"].nunique()
total_customers = df["customer_id"].nunique()
total_units = df["quantity"].sum()
total_revenue = df["sales_amount"].sum()
average_order_value = (
    total_revenue / total_orders if total_orders else 0
)
average_rating = df["rating"].mean()
return_rate = (
    df["returned"].astype(str).str.strip().str.lower()
      .eq("yes").mean() * 100
)

print("\n--- KEY PERFORMANCE INDICATORS ---")
print(f"Total orders: {total_orders:,}")
print(f"Unique customers: {total_customers:,}")
print(f"Units sold: {total_units:,.0f}")
print(f"Total revenue: ₹{total_revenue:,.2f}")
print(f"Average order value: ₹{average_order_value:,.2f}")
print(f"Average rating: {average_rating:.2f}/5")
print(f"Return rate: {return_rate:.2f}%")

# 4. Sales by product category
category_sales = (
    df.groupby("category")["sales_amount"]
      .sum()
      .sort_values(ascending=False)
)

print("\n--- REVENUE BY CATEGORY ---")
print(category_sales)

# 5. Sales by region
region_sales = (
    df.groupby("region")["sales_amount"]
      .sum()
      .sort_values(ascending=False)
)

print("\n--- REVENUE BY REGION ---")
print(region_sales)

# 6. Monthly revenue trend
df["month"] = df["order_date"].dt.to_period("M")

monthly_sales = (
    df.groupby("month")["sales_amount"]
      .sum()
)

print("\n--- MONTHLY REVENUE ---")
print(monthly_sales)

# 7. Payment method analysis
payment_analysis = df.groupby("payment_method").agg(
    orders=("order_id", "nunique"),
    revenue=("sales_amount", "sum"),
    average_rating=("rating", "mean")
).sort_values("revenue", ascending=False)

print("\n--- PAYMENT ANALYSIS ---")
print(payment_analysis)

# 8. Descriptive statistics using statistics module
revenue_values = df["sales_amount"].dropna().tolist()

print("\n--- DESCRIPTIVE STATISTICS ---")
if revenue_values:
    print("Mean:", statistics.mean(revenue_values))
    print("Median:", statistics.median(revenue_values))
    print("Standard deviation:",
          statistics.stdev(revenue_values)
          if len(revenue_values) > 1 else 0)

    try:
        print("Mode:", statistics.mode(revenue_values))
    except statistics.StatisticsError:
        print("Mode: No unique mode")

# 9. NumPy: correlation and simple linear regression
valid = df[["quantity", "sales_amount"]].dropna()
x = valid["quantity"].to_numpy(dtype=float)
y = valid["sales_amount"].to_numpy(dtype=float)

print("\n--- NUMPY ANALYSIS ---")
if len(x) >= 2 and np.std(x) > 0 and np.std(y) > 0:
    correlation = np.corrcoef(x, y)[0, 1]
    slope, intercept = np.polyfit(x, y, 1)

    print(f"Correlation (quantity vs sales): {correlation:.3f}")
    print(f"Regression slope: {slope:.2f}")
    print(f"Regression intercept: {intercept:.2f}")
else:
    print("Not enough variation for correlation/regression.")

# 10. Hypothesis test: do two payment groups have
# different average order-level revenue?
order_level = (
    df.groupby(["order_id", "payment_method"], as_index=False)
      .agg(order_revenue=("sales_amount", "sum"))
)

payment_totals = (
    order_level.groupby("payment_method")["order_revenue"]
      .sum()
      .sort_values(ascending=False)
)

if len(payment_totals) >= 2:
    group_names = list(payment_totals.index[:2])

    group_a = order_level.loc[
        order_level["payment_method"] == group_names[0],
        "order_revenue"
    ].dropna()

    group_b = order_level.loc[
        order_level["payment_method"] == group_names[1],
        "order_revenue"
    ].dropna()

    if len(group_a) >= 2 and len(group_b) >= 2:
        test = stats.ttest_ind(
            group_a, group_b, equal_var=False
        )

        print("\n--- WELCH'S T-TEST ---")
        print("Group A:", group_names[0], "n =", len(group_a))
        print("Group B:", group_names[1], "n =", len(group_b))
        print(f"t-statistic: {test.statistic:.3f}")
        print(f"p-value: {test.pvalue:.5f}")

        if test.pvalue < 0.05:
            print("Statistically significant difference at 5%.")
        else:
            print("No statistically significant difference at 5%.")

# 11. Visualizations
category_sales.plot(kind="bar", title="Revenue by Category")
plt.xlabel("Category")
plt.ylabel("Revenue (₹)")
plt.tight_layout()
plt.show()

monthly_sales.index = monthly_sales.index.astype(str)
monthly_sales.plot(kind="line", marker="o",
                   title="Monthly Revenue Trend")
plt.xlabel("Month")
plt.ylabel("Revenue (₹)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# 12. Export analysis results
category_sales.to_csv(
    r"Data/Processed/revenue_by_category.csv",
    header=["revenue"]
)

region_sales.to_csv(
    r"Data/Processed/revenue_by_region.csv",
    header=["revenue"]
)

monthly_sales.to_csv(
    r"Data/Processed/monthly_revenue.csv",
    header=["revenue"]
)

print("\nAnalysis completed successfully!")
