import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Set plot style
sns.set(style="whitegrid"

# Required columns
required_columns = ['Date', 'Total Sales', 'Net Revenue', 'Product', 'Units', 'Region', 
                    'City', 'Shipping Fee', 'Order Status', 'Discount (%)', 'Payment Method', 'Day']

# Load the Excel file
file_path = "Data set.xlsx"
df = pd.read_excel(file_path, sheet_name="DATA")

# Validate required columns
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise KeyError(f"Missing required columns: {missing_columns}")

# Data Cleaning
# 1. Remove duplicates
df = df.drop_duplicates()

# 2. Handle missing values (drop rows with missing critical columns)
df = df.dropna(subset=['Date', 'Net Revenue', 'Units'])

# 3. Convert date column and drop invalid dates
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df = df.dropna(subset=['Date'])

# 4. Derive 'Day' from 'Date'
df['Day'] = df['Date'].dt.day_name()

# 5. Standardize categorical columns
categorical_cols = ['Product', 'Region', 'City', 'Order Status', 'Payment Method']
for col in categorical_cols:
    df[col] = df[col].str.title()

# 6. Fill remaining missing values
df.fillna({'Total Sales': 0, 'Shipping Fee': 0, 'Discount (%)': 0,
           'Product': 'Unknown', 'Region': 'Unknown', 'City': 'Unknown',
           'Order Status': 'Unknown', 'Payment Method': 'Unknown'}, inplace=True)

# 7. Validate numeric columns (ensure non-negative)
numeric_cols = ['Total Sales', 'Net Revenue', 'Units', 'Shipping Fee', 'Discount (%)']
for col in numeric_cols:
    df[col] = df[col].clip(lower=0)

# 8. Handle outliers using IQR method
for col in ['Total Sales', 'Net Revenue', 'Units', 'Shipping Fee']:
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    df[col] = df[col].clip(lower=lower_bound, upper=upper_bound)

# Show column names
print("\nColumn names:\n", df.columns.tolist())

# Check for null values
print("\nNull values in each column:\n", df.isnull().sum())

# Data exploration
print("HEAD:")
print(df.head())
print("Describe:")
print(df.describe())
print("INFO:")
df.info()

# 1. Total Sales & Net Revenue Over Time
daily_sales = df.groupby('Date')[['Total Sales', 'Net Revenue']].sum()
plt.figure(figsize=(12, 6))
daily_sales.plot(title="Daily Total Sales vs Net Revenue")
plt.ylabel("Amount (₹)")
plt.xlabel("Date")
plt.tight_layout()
plt.savefig("daily_sales.png", dpi=300, bbox_inches='tight')
plt.show()

# 2. Top 10 Best-Selling Products (Units)
top_products = df.groupby('Product')['Units'].sum().nlargest(10)
plt.figure(figsize=(10, 6))
sns.barplot(x=top_products.values, y=top_products.index, palette="viridis")
plt.title("Top 10 Best-Selling Products (Units Sold)")
plt.xlabel("Units Sold")
plt.tight_layout()
plt.savefig("top_products_units.png", dpi=300, bbox_inches='tight')
plt.show()

# 3. Top 10 Products by Net Revenue
top_revenue_products = df.groupby('Product')['Net Revenue'].sum().nlargest(10)
plt.figure(figsize=(10, 6))
sns.barplot(x=top_revenue_products.values, y=top_revenue_products.index, palette="magma")
plt.title("Top 10 Products by Net Revenue")
plt.xlabel("Net Revenue (₹)")
plt.tight_layout()
plt.savefig("top_products_revenue.png", dpi=300, bbox_inches='tight')
plt.show()

# 4. Revenue by Region
region_sales = df.groupby('Region')['Net Revenue'].sum().sort_values(ascending=False)
plt.figure(figsize=(8, 5))
sns.barplot(x=region_sales.index, y=region_sales.values, palette="pastel")
plt.title("Net Revenue by Region")
plt.ylabel("Revenue (₹)")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig("region_sales.png", dpi=300, bbox_inches='tight')
plt.show()

# 5. City-wise Shipping Fee vs Net Revenue (Top 10 Cities)
city_group = df.groupby('City')[['Shipping Fee', 'Net Revenue']].sum().nlargest(10, 'Net Revenue')
city_group.plot(kind='bar', figsize=(12, 6), title="Top Cities: Shipping Fee vs Net Revenue")
plt.ylabel("Amount (₹)")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig("city_shipping_revenue.png", dpi=300, bbox_inches='tight')
plt.show()

# 6. Order Status Distribution
plt.figure(figsize=(6, 6))
df['Order Status'].value_counts().plot.pie(autopct='%1.1f%%', startangle=90, colors=sns.color_palette("Set2"))
plt.title("Order Status Distribution")
plt.ylabel("")
plt.tight_layout()
plt.savefig("order_status.png", dpi=300, bbox_inches='tight')
plt.show()

# 7. Units Ordered per Transaction
plt.figure(figsize=(10, 5))
sns.histplot(df['Units'], bins=30, kde=True, color='skyblue')
plt.title("Units Ordered per Transaction")
plt.xlabel("Units")
plt.tight_layout()
plt.savefig("units_per_transaction.png", dpi=300, bbox_inches='tight')
plt.show()

# 8. Preferred Payment Methods
plt.figure(figsize=(8, 5))
sns.countplot(data=df, x='Payment Method', order=df['Payment Method'].value_counts().index, palette="cubehelix")
plt.title("Payment Method Preference")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig("payment_methods.png", dpi=300, bbox_inches='tight')
plt.show()

# 9. Sales by Day of the Week
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
day_sales = df.groupby('Day')['Net Revenue'].sum().reindex(day_order)
plt.figure(figsize=(10, 6))
sns.barplot(x=day_sales.index, y=day_sales.values, palette="coolwarm")
plt.title("Net Revenue by Day of the Week")
plt.ylabel("Revenue (₹)")
plt.tight_layout()
plt.savefig("day_sales.png", dpi=300, bbox_inches='tight')
plt.show()

# Close all plots
plt.close('all')
