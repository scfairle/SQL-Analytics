
import pandas as pd
import matplotlib.pyplot as plt

# Load Excel file and select SQL sheet
file_path = "Analytics_Quiz (3) (1) (1).xlsx"
df = pd.read_excel(file_path, sheet_name="SQL", skiprows=1)

# Extract input table
input_df = df.iloc[3:, 0:3].copy()
input_df.columns = ["Order Date", "SKU", "Units Sold"]
input_df = input_df.dropna()

# Convert types
input_df["Order Date"] = pd.to_datetime(input_df["Order Date"])
input_df["Units Sold"] = pd.to_numeric(input_df["Units Sold"])
input_df["Month"] = input_df["Order Date"].dt.strftime("%b")

# Pivot table by month
monthly_sales = input_df.pivot_table(
    index="SKU",
    columns="Month",
    values="Units Sold",
    aggfunc="sum",
    fill_value=0
)

# Reorder months and add Q1 Total
monthly_sales = monthly_sales[["Jan", "Feb", "Mar"]]
monthly_sales["Q1 Unit Sales"] = monthly_sales.sum(axis=1)

# Total sales and shares
monthly_totals = monthly_sales[["Jan", "Feb", "Mar"]].sum()
q1_total = monthly_sales["Q1 Unit Sales"].sum()
monthly_sales["Jan Share"] = monthly_sales["Jan"] / monthly_totals["Jan"]
monthly_sales["Feb Share"] = monthly_sales["Feb"] / monthly_totals["Feb"]
monthly_sales["Mar Share"] = monthly_sales["Mar"] / monthly_totals["Mar"]
monthly_sales["Q1 Share"] = monthly_sales["Q1 Unit Sales"] / q1_total

# Ranking
monthly_sales["Q1 Rank"] = monthly_sales["Q1 Unit Sales"].rank(ascending=False).astype(int)
monthly_sales = monthly_sales.round(3)

# Save to CSV (optional)
monthly_sales.to_csv("q1_sales_summary.csv")

# Visualization
top5 = monthly_sales.sort_values("Q1 Unit Sales", ascending=False).head(5)
plt.figure(figsize=(10, 6))
plt.bar(top5.index, top5["Q1 Unit Sales"], color="skyblue")
plt.title("Top 5 SKUs by Q1 Unit Sales")
plt.xlabel("SKU")
plt.ylabel("Units Sold")
plt.grid(axis='y')
plt.tight_layout()
plt.savefig("top5_q1_sales.png")
plt.show()
