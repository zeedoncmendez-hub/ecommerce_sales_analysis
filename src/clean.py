
import pandas as pd

# 1. Load the raw dataset
df = pd.read_csv(r"C:\Users\shobi\OneDrive\Desktop\project\E-Commerce_sales_analysis\Data\Raw\ecommerce_sales_raw.csv")

print("Original dataset shape:", df.shape)

# 2. Remove duplicate rows
df = df.drop_duplicates()

# 3. Clean column names
df.columns = df.columns.str.strip().str.lower()

# 4. Clean text columns
text_columns = [
    "gender", "city", "region", "category",
    "sub_category", "payment_method", "returned"
]

for col in text_columns:
    df[col] = df[col].astype("string").str.strip()

# Standardize text capitalization
df["gender"] = df["gender"].str.title()
df["city"] = df["city"].str.title()
df["region"] = df["region"].str.title()
df["category"] = df["category"].str.title()
df["sub_category"] = df["sub_category"].str.title()
df["payment_method"] = df["payment_method"].str.title()

# Standardize known category variations
df["category"] = df["category"].replace({
    "Electronics": "Electronics",
    "Home & Kitchen": "Home & Kitchen"
})

df["payment_method"] = df["payment_method"].replace({
    "Upi": "UPI",
    "Cash On Delivery": "Cash on Delivery",
    "Credit Card": "Credit Card",
    "Debit Card": "Debit Card",
    "Net Banking": "Net Banking",
    "Wallet": "Wallet"
})

df["returned"] = df["returned"].str.strip().str.lower()
df["returned"] = df["returned"].replace({
    "yes": "Yes",
    "no": "No"
})

# 5. Convert columns to appropriate data types
df["order_date"] = pd.to_datetime(
    df["order_date"], errors="coerce"
)

numeric_columns = [
    "customer_age", "unit_price", "quantity",
    "discount_pct", "delivery_days", "rating"
]

for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# 6. Replace invalid values with missing values
df.loc[
    ~df["customer_age"].between(18, 100),
    "customer_age"
] = pd.NA

df.loc[df["unit_price"] <= 0, "unit_price"] = pd.NA

df.loc[df["quantity"] <= 0, "quantity"] = pd.NA

df.loc[
    ~df["discount_pct"].between(0, 100),
    "discount_pct"
] = pd.NA

df.loc[df["delivery_days"] <= 0, "delivery_days"] = pd.NA

df.loc[~df["rating"].between(1, 5), "rating"] = pd.NA

# 7. Handle missing values
# Use median for numerical columns
for col in ["customer_age", "unit_price",
            "quantity", "delivery_days", "rating"]:
    df[col] = df[col].fillna(df[col].median())

# Use mode for categorical columns
for col in ["gender", "city", "region", "category",
            "sub_category", "payment_method", "returned"]:
    mode_value = df[col].mode(dropna=True)

    if not mode_value.empty:
        df[col] = df[col].fillna(mode_value.iloc[0])

# 8. Remove rows with invalid essential information
df = df.dropna(subset=["order_id", "order_date", "customer_id"])

# 9. Create a sales amount column after discount
df["sales_amount"] = (
    df["unit_price"]
    * df["quantity"]
    * (1 - df["discount_pct"] / 100)
).round(2)

# 10. Save the cleaned dataset
df.to_csv(r"C:\Users\shobi\OneDrive\Desktop\project\E-Commerce_sales_analysis\Data\Processed\ecommerce_sales_cleaned.csv", index=False)

# 11. Check the results
print("Cleaned dataset shape:", df.shape)
print("Remaining missing values:")
print(df.isnull().sum())

print("\nFirst 5 rows:")
print(df.head())

print("\nCleaned dataset saved successfully!")
