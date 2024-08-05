import pandas as pd

# Load the Excel files
reason_of_inclusion_file = "ReasonOfInclusion_EntryNumbers_REACH List.xlsx"
eu_reach_cleaned_file = "EU_REACH_2023_CLEANED.xlsx"

# Read the sheets into dataframes
reason_df = pd.read_excel(reason_of_inclusion_file)
eu_reach_df = pd.read_excel(eu_reach_cleaned_file)

# Create a dictionary to map Entry Numbers to Conditions of Restriction
entry_conditions_map = pd.Series(
    reason_df["Conditions of restriction"].values,
    index=reason_df["Entry Number"]
).to_dict()


# Define a function to map conditions based on entry numbers
def map_conditions(entry_number):
    return entry_conditions_map.get(entry_number, "")


# Update the EU_REACH_2023_CLEANED dataframe with conditions
eu_reach_df["Conditions of restriction"] = eu_reach_df["Entry Number"].map(map_conditions)

# Save the updated dataframe back to a new Excel file
updated_file = "EU_REACH_2023_CLEANED_UPDATED.xlsx"
eu_reach_df.to_excel(updated_file, index=False)

print(f"Updated file saved as {updated_file}")
