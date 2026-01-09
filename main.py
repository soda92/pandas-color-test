import pandas as pd
import numpy as np

# 1. Read your source Excel
df = pd.read_excel("source_data.xlsx")


# 2. Define your QC Logic
# This function receives a single value and returns a CSS string
def mark_invalid_data(val):
    """
    QC Rule:
    - If value is negative, mark RED.
    - If value is zero, mark YELLOW.
    - Otherwise, leave default.
    """
    if isinstance(val, (int, float)):
        if val < 0:
            return "background-color: #ff9999"  # Red
        elif val == 0:
            return "background-color: #ffffcc"  # Yellow
    return ""  # Default (no style)


# 3. Apply the style and save
# .map() applies the function element-wise
styler = df.style.map(mark_invalid_data)

# You can also use .apply() for column-wide logic (e.g. check if Col A > Col B)

# 4. Save to Excel
styler.to_excel("qc_report_styled.xlsx", index=False)
