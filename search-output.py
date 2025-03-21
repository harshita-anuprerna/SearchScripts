import json
import pandas as pd
import re

# Load the JSON file with the correct path
input_file = '/Users/harshita/Desktop/Work/CodeWork/Email templates/search-output.json'
with open(input_file, 'r') as file:
    data = json.load(file)

# List to store processed rows
rows = []

# Process each entry in the JSON data
for entry in data:
    row = {
        'id': entry['_id']['$oid'],
        'event-ID': entry['dataEventToken']['$oid'],
        'owner': entry['owner']['$oid'],
        'property-ID': entry['property']['$oid'],
        'url': entry['url'],
        'user-ID': entry['userId'],
        'ipAddress': entry['ipAddress'],
        'countryCode': entry['countryCode'],
        'region': entry['region'],
        'city': entry['city'],
        'browserName': entry['browserName'],
        'osName': entry['osName'],
        'deviceType': entry['deviceType'],
        'newUser': entry.get('newUser', ''),  # Use get() to avoid KeyError
        'newSession': entry.get('newSession', ''),  # Use get() to avoid KeyError
        'createdAt': entry['createdAt']['$date'],
        'updatedAt': entry['updatedAt']['$date'],
    }

    # Handle eventLogData dynamically - Only take keywords, skip numbers
    event_log_data = entry.get('eventLogData', {})
    for key, value in event_log_data.items():
        # Keep only string-based values (filter out numbers)
        if isinstance(value, str):
            row[f'EventLogData_{key}'] = value

    rows.append(row)

# Create a DataFrame
df = pd.DataFrame(rows)

# Fill missing columns dynamically if EventLogData has different keys
df = df.fillna('')

# -------- CLEAN ILLEGAL CHARACTERS --------
# Function to clean illegal characters from strings
def clean_illegal_characters(value):
    if isinstance(value, str):
        # Remove characters that are not allowed by Excel
        return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', value)
    return value

# Apply this function to all cells in the DataFrame
df = df.applymap(clean_illegal_characters)

# -------- SAVE TO EXCEL --------
# Save to Excel at the same location as the JSON file
output_file = '/Users/harshita/Desktop/Work/CodeWork/Email templates/search-output.xlsx'
df.to_excel(output_file, index=False)

print(f"✅ Data successfully written to '{output_file}'")
