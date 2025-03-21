import pandas as pd

# Load the previously generated Excel file
input_file = '/Users/harshita/Desktop/Work/CodeWork/Email templates/search-output.xlsx'
df = pd.read_excel(input_file)

# Filter columns to only include EventLogData_searchKey
search_key_column = 'EventLogData_searchKey'

# Check if the column exists in the DataFrame
if search_key_column not in df.columns:
    print(f"❌ Column '{search_key_column}' not found in the Excel file. Please check your data.")
    exit()

# Create a list to store all search keys
search_keys = []

# Extract, clean, and split values from EventLogData_searchKey
# Check for missing values, split by comma if multiple keywords
values = df[search_key_column].dropna().apply(lambda x: [item.strip().lower() for item in str(x).split(',')])

# Flatten the list and add to search_keys
for value_list in values:
    search_keys.extend(value_list)

# Remove empty strings and None values
search_keys = [key for key in search_keys if key != '' and key is not None]

# Create a DataFrame with search key frequency
search_key_df = pd.DataFrame({'SEARCH_KEY': search_keys})

# Group and count occurrences of each search key
search_key_df = search_key_df.groupby('SEARCH_KEY').size().reset_index(name='FREQUENCY')

# Sort by frequency in descending order
search_key_df = search_key_df.sort_values(by='FREQUENCY', ascending=False)

# Save the search key frequency to a new Excel file
output_file = '/Users/harshita/Desktop/Work/CodeWork/Email templates/keyword-frequency.xlsx'
search_key_df.to_excel(output_file, index=False)

print(f"✅ Search key frequency Excel generated successfully at: {output_file}")
