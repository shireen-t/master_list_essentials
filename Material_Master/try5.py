import pandas as pd

# Read the Excel file
file_path = 'MasterList_Viridium_002.xlsx'  # Replace with the path to your Excel file
df = pd.read_excel(file_path)

# Initialize an empty dictionary to store CAS numbers and their corresponding sources
cas_data = {}

# Iterate through the DataFrame rows
for index, row in df.iterrows():
    source = row['SOURCE']
    cas_no = row['CASNumber']
    name = row['MaterialName']  # Assuming there's a 'NAME' column in the source file
    msds = str(row['MSDSAvailability']).strip()  # Strip any leading/trailing spaces from the MSDS value
    
    # Print debug information
    print(f"Processing row {index}: Source={source}, CAS No={cas_no}, MSDS={msds}")
    
    # Skip rows with 'NONE' CAS numbers
    if cas_no != 'NONE':
        if cas_no not in cas_data:
            cas_data[cas_no] = {
                'MaterialName': name,
                'OECD': 'NO',
                'CANDIDATE': 'NO',
                'REACH': 'NO',
                'FORD': 'NO',
                'TSCA': 'NO',
                'MSDSAvailability': 'NO'
            }
        if source in cas_data[cas_no]:
            cas_data[cas_no][source] = 'YES'
        # Check for MSDS condition
        if msds == 'YES':
            cas_data[cas_no]['MSDSAvailability'] = 'YES'

# Convert the dictionary to a DataFrame
result_df = pd.DataFrame.from_dict(cas_data, orient='index').reset_index()
result_df.rename(columns={'index': 'CAS'}, inplace=True)

# Write the result to a new Excel file
output_file_path = 'EDITED.xlsx'  # Replace with the desired output file path
result_df.to_excel(output_file_path, index=False)

print(f"Results have been saved to {output_file_path}")