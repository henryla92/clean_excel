import pandas as pd
import re
import sys
import json

# Check if filename is provided
if len(sys.argv) < 2:
    print("Please provide filename as a command-line argument.")
    sys.exit()

# Load the Excel file
filename = sys.argv[1]
df = pd.read_excel(filename)

# Define a function to apply the regex
def extract_before_last_underscore(s):
    if isinstance(s, str):
        # parse the string as JSON to get the list
        items = json.loads(s)

        # iterate over the list and apply the regex to each item
        new_items = []
        for item in items:
            # Check if last character is an underscore
            if item[-1] == '_':
                match = re.search(r'(.*)(?=_[^_]*_$)', item)
            else:
                match = re.search(r'(.*)(?=_[^_]*$)', item)

            if match:
                new_items.append(match.group(1))
            else:
                new_items.append(item)

        # join the items with a comma and return the result
        return ', '.join(new_items)

    else:
        return s

# Apply the function to each column
for col in df.columns:
    df[col] = df[col].apply(extract_before_last_underscore)

# Save the modified DataFrame back to Excel
df.to_excel("modified_" + filename, index=False)
