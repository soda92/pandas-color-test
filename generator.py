import pandas as pd
import numpy as np

# 1. Configuration
filename = "source_data.xlsx"
rows = 50

# 2. Generate Data
# We use a seed so the random numbers are the same every time you run this
np.random.seed(42)

data = {
    "Transaction_ID": range(1001, 1001 + rows),
    "Date": pd.date_range(start="2024-01-01", periods=rows, freq="D"),
    # Integers: Mix of negative (Invalid), zero (Warning), and positive (Valid)
    "Inventory_Diff": np.random.randint(-10, 50, size=rows),
    # Floats: Standard normal distribution centered around 0
    "Profit_Margin": np.random.uniform(-0.5, 0.5, size=rows).round(2),
    # Categorical: Random status strings
    "Status": np.random.choice(["Confirmed", "Pending", "Error"], size=rows),
}

df = pd.DataFrame(data)

# 3. Inject explicit edge cases to ensure we have specific test scenarios
# Force row 5 to have a Zero (to test the Yellow condition)
df.loc[5, "Inventory_Diff"] = 0
# Force row 10 to have a Negative (to test the Red condition)
df.loc[10, "Inventory_Diff"] = -5

# 4. Save to Excel
df.to_excel(filename, index=False)

print(f"Generated '{filename}' with {rows} rows.")
print("Sample of generated data:")
print(df.head(10))
