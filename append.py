import pandas as pd

# Source and target file paths
source_file = "nafdac_greenbook.xlsx"
target_file = "nafdac_greenbook-data.xlsx"
output_file = "nafdac_greenbook-data_merged.xlsx"  # You can overwrite if you want

# Load the data
source_df = pd.read_excel(source_file)
target_df = pd.read_excel(target_file)

# Select first 1441 rows from the source file
rows_to_add = source_df.head(1441)

# Append them at the TOP of the target data
merged_df = pd.concat([rows_to_add, target_df], ignore_index=True)

# Save to new Excel file
merged_df.to_excel(output_file, index=False)

print(f"✅ Done! Saved merged file as {output_file}")

