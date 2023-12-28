import pandas as pd

# Lees de twee xlsx-bestanden in
df1 = pd.read_excel("Asana_saved.xlsx")
df2 = pd.read_excel("calendly_export.xlsx")

# Definieer de columns_to_add
columns_to_add = ["Question 1", "Response 1", "Question 2", "Response 2"]

# Remove duplicates in df1 based on "Name" column
df1 = df1.drop_duplicates("Name")

# Remove duplicates in df2 based on "Name" column
df2 = df2.drop_duplicates("Name")

# Check for duplicate values in the "Name" column of df2
duplicates = df2["Name"].duplicated(keep=False)
if duplicates.any():
    raise ValueError("Duplicate values found in the 'Name' column of df2. Please resolve the duplicates.")

# Merge based on the "Name" column
merged_df = pd.merge(df1, df2[["Name"] + columns_to_add], on="Name", how="left")

# Replace "Emailadres" in df1 with the corresponding data from df2
df2_unique = df2.drop_duplicates("Name")  # Drop duplicates to make the index unique
df1["Emailadres"] = df1["Name"].map(df2_unique.set_index("Name")["Emailadres"])

# Update the columns_to_add in df1 with the values from merged_df
df1[columns_to_add] = merged_df[columns_to_add]

# Add rows from df2 that do not match with existing users in df1
new_rows = df2[~df2["Name"].isin(df1["Name"])]
result_df = pd.concat([df1, new_rows[["Name", "Emailadres"] + columns_to_add]], ignore_index=True)

# Add a new column "meeting Calendly" and set True for manipulated rows, False otherwise
result_df["meeting Calendly"] = result_df["Name"].isin(df2["Name"])

# Schrijf de DataFrame terug naar een xlsx-bestand
result_df.to_excel("result.xlsx", index=False)
