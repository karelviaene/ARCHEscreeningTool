import pandas as pd
import re

# Load the Excel files
df = pd.read_excel("EDscreener_results.xlsx")          # your main data
CASallpd = pd.read_excel("testCAS.xlsx")  # file containing the desired order

# Choose the column to sort by
selected_col = "CAS"

# Get desired order from CASallpd
CASall = CASallpd[selected_col].dropna().tolist()  # selected_col is chosen in the app
CASall = [re.sub(r'[^\d\-]', '', str(cas)) for cas in CASall]
order = CASall
print(order)

# Convert df[selected_col] to categorical with that order
df["Input"] = pd.Categorical(df["Input"], categories=order, ordered=True)

# Sort and reset index
df = df.sort_values(by="Input").reset_index(drop=True)

# Optionally save back to Excel
df.to_excel("sorted_data.xlsx", index=False)

print("Sorted dataframe saved to sorted_data.xlsx")
