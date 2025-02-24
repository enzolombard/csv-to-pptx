import pandas as pd

# Load the CSV file
df = pd.read_csv("test_sheet.csv")

# Drop completely empty rows
df.dropna(how='all', inplace=True)

# Drop completely empty columns
df.dropna(axis=1, how='all', inplace=True)

# Save the cleaned file
df.to_csv("cleaned_file.csv", index=False)
