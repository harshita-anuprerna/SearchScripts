import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel sheet
input_file = "/Users/harshita/Desktop/Work/CodeWork/Email templates/search-output.xlsx"  # Adjust path if needed
df = pd.read_excel(input_file)

# Ask user for the keyword to analyze
keyword = input("Enter the keyword to analyze: ").strip().lower()

# Find all EventLogData columns dynamically
event_cols = [col for col in df.columns if col.startswith("EventLogData")]

# Ensure that we work with a copy of the dataframe without modifying the original index
df_filtered = df.copy()

# Create a boolean mask to find matching rows
mask = df_filtered[event_cols].astype(str).apply(
    lambda row: any(keyword in item.strip().lower() for col in event_cols for item in str(row[col]).split(",") if pd.notna(row[col])),
    axis=1
)

# Apply mask correctly to select only matching rows
df_filtered = df_filtered.loc[mask].copy()  # Ensure copy to prevent SettingWithCopyWarning

# Convert 'createdAt' to datetime and extract the month-year for trends
df_filtered["createdAt"] = pd.to_datetime(df_filtered["createdAt"], errors="coerce")
df_filtered.dropna(subset=["createdAt"], inplace=True)  # Drop rows with invalid dates
df_filtered["month_year"] = df_filtered["createdAt"].dt.to_period("M")

# Count occurrences of the keyword per month
monthly_counts = df_filtered["month_year"].value_counts().sort_index()

# Plot the trend
plt.figure(figsize=(10, 6))
monthly_counts.plot(kind="bar", color="skyblue")
plt.title(f"Search Trend for '{keyword}' Over Time")
plt.xlabel("Month-Year")
plt.ylabel("Frequency")
plt.xticks(rotation=45)
plt.tight_layout()

# Save the graph as an image
output_file = "/Users/harshita/Desktop/Work/CodeWork/Email templates/keyword_trend.png"
plt.savefig(output_file)
print(f"Trend graph saved at: {output_file}")

plt.show()
