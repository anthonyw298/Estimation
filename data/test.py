import pandas as pd

# Google Sheet details
sheet_id = "1F4qGDCsr-5sC9eymsnlcYPua_u9BeFFA"
gid = "1267775724"

# Build the CSV export link
csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"

# Load the sheet
df = pd.read_csv(csv_url)

# Get the first 5 columns
df_first5 = df.iloc[:, :5]

# Print without the index
print(df_first5.to_string(index=False))
