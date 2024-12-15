import pandas as pd
import requests

# Configuration
EXCEL_FILE = 'data.xls'  # Update with your actual file
FORM_URL = 'https://docs.google.com/forms/u/0/d/e/1FAIpQLSdUCd3UWQ3VOgeg0ZzNeT-xzNawU8AJ7Xidml-w1vhfBcvBWQ/formResponse'

# Read and clean data
try:
    data = pd.read_excel(EXCEL_FILE)
    data.columns = data.columns.str.strip()  # Remove leading/trailing spaces
    print("Column names:", data.columns)
except Exception as e:
    print(f"Error reading the Excel file: {e}")
    exit()

# Map your form fields to Google Form field IDs
FIELD_MAP = {
    'Full Name': 'entry.1234567890',  # Replace with actual field ID
    'Contact Number': 'entry.0987654321',  # Replace with actual field ID
    'Email ID': 'entry.1122334455',  # Replace with actual field ID
    'Full Address': 'entry.5566778899',  # Replace with actual field ID
    'Pin Code': 'entry.6677889900',  # Replace with actual field ID
    'Date of Birth': 'entry.7788990011',  # Replace with actual field ID
    'Gender': 'entry.8899001122',  # Replace with actual field ID
    'Verification Code': 'entry.9900112233'  # Replace with actual field ID
}

# Check if all required columns exist
missing_columns = [col for col in FIELD_MAP.keys() if col not in data.columns]
if missing_columns:
    print(f"Missing columns in Excel: {missing_columns}")
    exit()

# Iterate through each row in the Excel file
for index, row in data.iterrows():
    try:
        # Prepare payload
        form_data = {}
        for col_name, field_id in FIELD_MAP.items():
            form_data[field_id] = row[col_name]

        # Submit the form
        response = requests.post(FORM_URL, data=form_data)
        if response.status_code == 200:
            print(f"Form #{index + 1} submitted successfully.")
        else:
            print(f"Error submitting form #{index + 1}: HTTP {response.status_code}")

    except KeyError as e:
        print(f"An error occurred for form #{index + 1}: Missing column {e}")
    except Exception as e:
        print(f"An error occurred for form #{index + 1}: {e}")

print("All forms processed.")
