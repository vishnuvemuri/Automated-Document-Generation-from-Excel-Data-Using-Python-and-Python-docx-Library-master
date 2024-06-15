Automated Document Generation from Excel Data Using Python and Python-docx Library

Overview

This project focuses on automating the generation of Word documents from Excel data using Python and the python-docx library. It streamlines the process of creating standardized documents by extracting data from Excel sheets and populating pre-defined templates, thereby saving time and reducing errors associated with manual document creation.

Features

Excel Data Extraction: Reads and processes data from Excel files.
Template-Based Document Generation: Uses predefined Word templates to ensure consistency.
Dynamic Content Insertion: Automatically inserts data into specific locations within the Word document.
Batch Processing: Generates multiple documents in a single run based on the rows of the Excel sheet.
Customizable Templates: Supports the use of customizable Word templates to fit various document formats.
Error Handling: Includes robust error handling to manage common issues like missing data or template errors.
Installation
Clone the Repository:

sh
Copy code
git clone https://github.com/vishnuvemuri/automated-doc-gen.git
cd automated-doc-gen
Create a Virtual Environment:

sh
Copy code
python -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
Install Dependencies:

sh
Copy code
pip install -r requirements.txt
Usage
Preparing the Excel Data
Format Your Excel Sheet: Ensure your Excel file is formatted correctly. Each row should contain data for one document, and the columns should correspond to placeholders in your Word template.
Save the Excel File: Place your Excel file in the data/ directory.
Setting Up the Word Template
Create a Template: Design a Word document with placeholders for the data. Use unique identifiers (e.g., {{placeholder}}) where data should be inserted.
Save the Template: Place your Word template in the templates/ directory.
Generating Documents
Run the Script:

sh
Copy code
python generate_documents.py --excel data/yourfile.xlsx --template templates/yourtemplate.docx --output output/
Check Output: The generated documents will be saved in the output/ directory.

Customization
Configuring Placeholders: Update the script to match the placeholders in your template with the columns in your Excel sheet.
Adjusting Formatting: Modify the template or the script to adjust the formatting as needed.
Project Structure
data/: Directory for storing input Excel files.
templates/: Directory for storing Word templates.
output/: Directory for storing generated Word documents.
scripts/: Directory for utility scripts.
generate_documents.py: Main script for generating documents.
requirements.txt: List of required Python packages.
Contributing
We welcome contributions from the community. To contribute, please follow these steps:

Fork the repository.
Create a new branch (git checkout -b feature-branch).
Make your changes.
Commit your changes (git commit -am 'Add new feature').
Push to the branch (git push origin feature-branch).
Create a new Pull Request.
License
This project is licensed under the MIT License. See the LICENSE file for more details.
