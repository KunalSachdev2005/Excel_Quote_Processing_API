# Quote Extraction Automation API
This repository contains the code and documentation for the Quote Extraction Automation API developed during my internship at HDFC ERGO from April 2022 to June 2022. The project aimed to automate the validation, extraction, and storage of data from insurance quote Excel templates, improving the efficiency and accuracy of the policy issuance process.

## Project Overview
### Objective
The primary objective of this project was to create an API that automates the process of validating, extracting, and storing data from insurance quote Excel files into a database, thereby eliminating manual data entry and speeding up the policy issuance process.

### Current Process
![Current Process](https://github.com/KunalSachdev2005/Quote_Extraction_Automation_API/blob/main/media/Current_Process.png)

### Challenges with Current Process
1. Inconsistent Excel quote templates due to:
   - Changes in formulas by salespersons.
   - Use of obsolete templates.
2. Time-consuming & error-prone manual data entry.

### Modified Process
![Modified Process](https://github.com/KunalSachdev2005/Quote_Extraction_Automation_API/blob/main/media/Modified_Process.png)


### Advantages of Modified Process

1. **4 Factor Validation**:
   - File format check (.xlsx).
   - Unique hidden encrypted value check in a random cell.
   - Template format check using 5 constant labels.
   - Sheet properties check (creator and title).

2. **Automated Data Entry**:
   - Saves time (from 2 hours to 2 minutes per proposal).
   - Eliminates data entry errors.

3. **Managing Redundancy in Lead ID**

To handle redundancy in lead IDs:

An "isactive" column is introduced in each database table.
Upon data upload, the API checks for existing lead IDs and updates the "isactive" status accordingly, ensuring the most recent activity is marked as active.
Products
The API handles the following insurance products:

Standard Fire & Special Perils Policy (SFSP)
Burglary Insurance
Plate Glass Insurance
For each product, the API processes sections such as Customer Details, Main Policies, Add-On Covers, Terms & Conditions, Deductibles, Warranties, Clauses & Conditions, Exclusions & Subjectivities, and Supplementary Clauses & Conditions.

API Code Flow
File Upload: The user uploads the quote file.
Quote Format Validation: The API validates the file format and constant sheet values.
Quote Extraction: The API extracts data and uploads it to the database.
Demo
The repository includes a demo showcasing the front-end for file upload and the resulting output message.

Software Specifications
Python: 3.8.8
Database: Oracle 19c
Libraries:
Flask: 2.1.2
Pandas: 1.4.2
NumPy: 1.22.4
os: 0.6.3
sqlalchemy: 1.3.24
openpyxl: 3.0.9
Werkzeug: 2.1.2
cx_Oracle: 8.3.0
