import pandas as pd

# Get the file path from the user
file_path = input("Enter the file path of the desired .XLSX file: ")

# Get the buyer/seller value from the user
buyer_seller = input("Enter B for Buyers, S for Sellers: ")

# Get the date cutoff from the user
date_cutoff = input("Enter the date cutoff: ")

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Calculate the email string based off the user input
email_string = ''
if buyer_seller == 'B':
    email_string = 'Buyer Email'
else:
    email_string = 'Seller Email'

# Filter out rows with empty email addresses
df = df[df[email_string].notna()]

# Filter out rows where the "Date" column is less than the date cutoff based off the user input
filtered_df = df.loc[df['Date'] >= date_cutoff, :].copy()

# Sort the DataFrame by the "Date" column in ascending order (useful for comparing in order to test for accuracy)
filtered_df.sort_values(by='Date', inplace=True)

# Save the "Email" column of the filtered DataFrame to a new text file, separated by commas
with open('emails.txt', 'w') as f:
    f.write(','.join(filtered_df[email_string]))
