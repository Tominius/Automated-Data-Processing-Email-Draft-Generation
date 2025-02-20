# Automated-Data-Processing-Email-Draft-Generation

This Python script automates the workflow of processing Excel files, extracting client-specific data, and generating email drafts (.eml) with attachments.

# Features
✅ Excel to CSV Conversion – Cleans and converts .xlsx files to .csv, removing empty rows and columns.

✅ Client-Based Excel File Generation – Extracts specific columns from a CSV file and generates separate Excel files per client.

✅ Automated Email Draft Creation – Reads client emails from a CSV file and generates .eml drafts with pre-written messages and attached reports.


# Workflow
Convert Excel to CSV --> Reads an .xlsx file, removes empty data, and saves it as a .csv file.

Create Individual Excel Files for Clients --> Extracts relevant data from the CSV and generates personalized Excel reports for each client.

Generate Email Drafts (.eml) --> Uses client email data to create drafts with attached reports, ready to be sent.

# Requirements
Python 3.x
pandas
os
csv
email.message
mimetypes

# 📌 How to Use
1️⃣ Convert an Excel file to CSV:

Place your .xlsx file in the appropriate directory and modify the script to specify its path.
Run the script to generate a cleaned .csv file.

2️⃣ Generate client-specific Excel files:

Ensure that the CSV file contains a column for client names.
Run the script to create separate .xlsx files for each client.

3️⃣ Create email drafts (.eml) with attachments:

Provide a clientes_mails.csv file that maps client names to email addresses.
Run the script to generate .eml drafts, ready for sending.

# 🏗 Project Structure

📂 AdminTool/

 ├── 📂 Excels/         # Generated client-specific Excel files
 
 ├── 📂 Mails/          # Saved email drafts (.eml)
 
 ├── 📄 clientes_mails.csv  # Mapping of clients to email addresses
 
 ├── 📄 script.py       # Main Python script

# 🛠 Customization

Modify the column positions (posiciones) to extract relevant data from your CSV.
Update the email template in the correos function to personalize messages.

# 🚀 Contributing

Feel free to submit issues, fork the repository, or suggest improvements!

 
