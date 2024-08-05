import pandas as pd

# Load the Excel files
eu_reach_updated_file = "EU_REACH_2023_CLEANED_UPDATED.xlsx"
master_list_file = "Master_List.xlsx"

# Read the sheets into dataframes
eu_reach_df = pd.read_excel(eu_reach_updated_file)
master_list_df = pd.read_excel(master_list_file)

# Create a dictionary to map CAS Numbers to Conditions of Restriction
cas_conditions_map = pd.Series(
    eu_reach_df["Conditions of restriction"].values,
    index=eu_reach_df["CAS No."]
).to_dict()


# Define a function to map conditions based on CAS numbers
def map_conditions(cas_number):
    return cas_conditions_map.get(cas_number, "")


# Update the Master_List dataframe with conditions
master_list_df["Conditions of restriction"] = master_list_df["CASNumber"].map(map_conditions)

# Save the updated dataframe back to a new Excel file
updated_master_list_file = "Master_List_UPDATED.xlsx"
master_list_df.to_excel(updated_master_list_file, index=False)

print(f"Updated file saved as {updated_master_list_file}")
