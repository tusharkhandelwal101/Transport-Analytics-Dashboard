# Step 1: Import libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Step 2: Load the Excel dataset
df = pd.read_excel("Transport.xlsx")  

# # Step 3: Show first 5 rows
# print(df.head())

# # Step 4: Check data info
# print(df.info())

# # Step 5: Describe numerical data
# print(df.describe())

# # Step 6: Check missing values
# print(df.isnull().sum())

# Step 7: Revenue distribution plot
# plt.figure(figsize=(8, 5))
# sns.histplot(df["Revenue"], bins=30, kde=True)
# plt.title("Distribution of Revenue")
# plt.xlabel("Revenue")
# plt.ylabel("Frequency")
# plt.show()

# plt.figure(figsize=(10,6))
# sns.histplot(df['ShippingCost'], kde=True, color='lightcoral')
# plt.title("Distribution of Shipping Cost")
# plt.xlabel("Shipping Cost")
# plt.ylabel("Frequency")
# plt.show()

# plt.figure(figsize=(10,6))
# sns.scatterplot(data=df, x="TotalMiles", y="Revenue", hue="TripType")
# plt.title("Total Miles vs Revenue")
# plt.xlabel("Total Miles")
# plt.ylabel("Revenue")
# plt.show()

# plt.figure(figsize=(10,6))
# sns.countplot(data=df, x="ShipDays", palette="coolwarm")
# plt.title("Number of Trips per ShipDay")
# plt.xlabel("Ship Days")
# plt.ylabel("Count")
# plt.show()

# plt.figure(figsize=(12,6))
# sns.boxplot(data=df, x="OriginState", y="Revenue")
# plt.title("Revenue Distribution by Origin State")
# plt.xticks(rotation=45)
# plt.show()

# Clean and processing data
df = df.dropna()  # Or fillna() if needed
df = df[df['Revenue'] > -2000]  # Remove extreme outliers

#  a) Revenue by OriginState
rev_state = df.groupby('OriginState')['Revenue'].agg(['mean', 'sum', 'count']).reset_index()
rev_state.columns = ['OriginState', 'AvgRevenue', 'TotalRevenue', 'TotalTrips']

# 🛣️ b) Miles Efficiency by State
efficiency = df.groupby('OriginState').agg({
    'LoadedMiles': 'mean',
    'TotalMiles': 'mean',
    'Revenue': 'mean'
}).reset_index()

# 📦 c) Shipping Days Analysis
days_analysis = df.groupby('ShipDays').agg({
    'Revenue': ['mean', 'count'],
    'ShippingCost': 'mean'
}).reset_index()
days_analysis.columns = ['ShipDays', 'AvgRevenue', 'TotalTrips', 'AvgCost']

# 📍 d) Origin-Destination Pair Performance
pair_perf = df.groupby(['OriginCity', 'DestinationCity']).agg({
    'Revenue': 'sum',
    'ShippingCost': 'sum',
    'TotalMiles': 'mean'
}).reset_index()

# Export to Excel / CSV
with pd.ExcelWriter("EDA_Summary_Report.xlsx") as writer:
    rev_state.to_excel(writer, sheet_name="Revenue_By_State", index=False)
    efficiency.to_excel(writer, sheet_name="Efficiency_By_State", index=False)
    days_analysis.to_excel(writer, sheet_name="Shipping_Days_Analysis", index=False)
    pair_perf.to_excel(writer, sheet_name="OD_Performance", index=False)

