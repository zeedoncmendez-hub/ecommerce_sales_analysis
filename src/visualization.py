
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# ---------------------------------
# Paths and style
# ---------------------------------

PROJECT_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = PROJECT_DIR / "Data" / "Processed" / "ecommerce_sales_cleaned.csv"
OUTPUT_DIR = PROJECT_DIR / "outputs" / "charts"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

df = pd.read_csv(DATA_PATH)
df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")

if "sales_amount" not in df.columns:
    df["sales_amount"] = (
        df["unit_price"]
        * df["quantity"]
        * (1 - df["discount_pct"] / 100)
    )

df = df.dropna(subset=["order_date", "sales_amount"])
df["sales_amount"] = pd.to_numeric(
    df["sales_amount"], errors="coerce"
)
df = df.dropna(subset=["sales_amount"])

# One consistent colour per chart
SINGLE_COLOR = "#3973AC"
SECONDARY_COLOR = "#D97745"

plt.rcParams.update({
    "figure.figsize": (10, 6),
    "figure.dpi": 120,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.alpha": 0.20,
    "grid.linestyle": "--",
    "axes.axisbelow": True,
    "font.size": 10,
})


def save_chart(filename):
    """Save the current chart and close its figure."""
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, bbox_inches="tight")
    plt.close()


def label_bars(ax, horizontal=False, money=False):
    """Place values directly on or beside bars."""
    for container in ax.containers:
        labels = []
        for bar in container:
            value = (
                bar.get_width() if horizontal
                else bar.get_height()
            )
            if money:
                labels.append(f"₹{value:,.0f}")
            else:
                labels.append(f"{value:,.0f}")

        if horizontal:
            ax.bar_label(
                container, labels=labels, padding=4, fontsize=9
            )
        else:
            ax.bar_label(
                container, labels=labels, padding=3, fontsize=9
            )


# ---------------------------------
# 1. Revenue by product category
# ---------------------------------

category_sales = (
    df.groupby("category")["sales_amount"]
    .sum()
    .sort_values(ascending=True)
)

fig, ax = plt.subplots()
category_sales.plot(kind="barh", ax=ax, color=SINGLE_COLOR)

ax.set_title("Revenue contribution varies across product categories",
             loc="left", fontweight="bold")
ax.set_xlabel("Revenue (₹)")
ax.set_ylabel("Product category")
ax.xaxis.set_major_formatter(
    mtick.StrMethodFormatter("₹{x:,.0f}")
)
ax.grid(axis="y", visible=False)

label_bars(ax, horizontal=True, money=True)
save_chart("01_category_revenue.png")


# ---------------------------------
# 2. Monthly revenue trend
# ---------------------------------

monthly_sales = (
    df.set_index("order_date")["sales_amount"]
    .resample("MS")
    .sum()
)

fig, ax = plt.subplots()

ax.plot(
    monthly_sales.index,
    monthly_sales.values,
    color=SINGLE_COLOR,
    linewidth=2.5,
    marker="o",
    markersize=5,
)

ax.set_title("Monthly revenue shows how sales change over time",
             loc="left", fontweight="bold")
ax.set_xlabel("Month")
ax.set_ylabel("Revenue (₹)")
ax.yaxis.set_major_formatter(
    mtick.StrMethodFormatter("₹{x:,.0f}")
)
ax.tick_params(axis="x", rotation=35)

# Label each monthly point
for date, value in monthly_sales.items():
    ax.annotate(
        f"₹{value:,.0f}",
        (date, value),
        xytext=(0, 8),
        textcoords="offset points",
        ha="center",
        fontsize=8,
    )

save_chart("02_monthly_revenue.png")


# ---------------------------------
# 3. Order volume by payment method
# ---------------------------------

payment_orders = (
    df.groupby("payment_method")["order_id"]
    .nunique()
    .sort_values(ascending=True)
)

fig, ax = plt.subplots()
payment_orders.plot(kind="barh", ax=ax, color=SINGLE_COLOR)

ax.set_title("Order volume differs by payment method",
             loc="left", fontweight="bold")
ax.set_xlabel("Unique orders")
ax.set_ylabel("Payment method")
ax.grid(axis="y", visible=False)

label_bars(ax, horizontal=True)
save_chart("03_payment_method_orders.png")


# ---------------------------------
# 4. Average rating by category
# ---------------------------------

if "rating" in df.columns:
    ratings = df.copy()
    ratings["rating"] = pd.to_numeric(
        ratings["rating"], errors="coerce"
    )
    ratings = ratings.dropna(subset=["rating"])

    category_rating = (
        ratings.groupby("category")["rating"]
        .mean()
        .sort_values(ascending=True)
    )

    fig, ax = plt.subplots()
    category_rating.plot(
        kind="barh", ax=ax, color=SINGLE_COLOR
    )

    ax.set_title("Average customer rating varies by category",
                 loc="left", fontweight="bold")
    ax.set_xlabel("Average rating (out of 5)")
    ax.set_ylabel("Product category")
    ax.set_xlim(0, 5)
    ax.grid(axis="y", visible=False)

    for container in ax.containers:
        ax.bar_label(
            container,
            labels=[f"{v:.2f}" for v in category_rating.values],
            padding=3,
            fontsize=9,
        )

    save_chart("04_category_ratings.png")


# ---------------------------------
# 5. Distribution of order revenue
# ---------------------------------

fig, ax = plt.subplots()

ax.hist(
    df["sales_amount"],
    bins=30,
    color=SINGLE_COLOR,
    edgecolor="white",
)

median_revenue = df["sales_amount"].median()
ax.axvline(
    median_revenue,
    color=SECONDARY_COLOR,
    linewidth=2,
    linestyle="--",
    label=f"Median: ₹{median_revenue:,.0f}",
)

ax.set_title("Order revenue distribution reveals typical order sizes",
             loc="left", fontweight="bold")
ax.set_xlabel("Revenue per order (₹)")
ax.set_ylabel("Number of records")
ax.legend(frameon=False)

save_chart("05_revenue_distribution.png")


# ---------------------------------
# 6. Return rate by category
# ---------------------------------

if "returned" in df.columns:
    returns = df.copy()
    returns["returned"] = (
        returns["returned"]
        .astype(str)
        .str.strip()
        .str.lower()
        .map({
            "true": True,
            "false": False,
            "1": True,
            "0": False,
            "yes": True,
            "no": False,
        })
    )

    returns = returns.dropna(subset=["returned"])

    category_return_rate = (
        returns.groupby("category")["returned"]
        .mean()
        .mul(100)
        .sort_values(ascending=True)
    )

    fig, ax = plt.subplots()
    category_return_rate.plot(
        kind="barh", ax=ax, color=SINGLE_COLOR
    )

    ax.set_title("Return rates differ across product categories",
                 loc="left", fontweight="bold")
    ax.set_xlabel("Returned records (%)")
    ax.set_ylabel("Product category")
    ax.xaxis.set_major_formatter(
        mtick.PercentFormatter(xmax=100)
    )
    ax.grid(axis="y", visible=False)

    for container in ax.containers:
        ax.bar_label(
            container,
            labels=[f"{v:.1f}%" for v in category_return_rate.values],
            padding=3,
            fontsize=9,
        )

    save_chart("06_category_return_rate.png")


print("Visualisation completed.")
print(f"Charts saved to: {OUTPUT_DIR}")
