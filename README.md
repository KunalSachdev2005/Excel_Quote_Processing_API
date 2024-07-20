# Quote Extraction Automation API
This repository contains the code and documentation for the Quote Extraction Automation API developed during my internship at HDFC ERGO from April 2022 to June 2022. The project aimed to automate the validation, extraction, and storage of data from insurance quote Excel templates, improving the efficiency and accuracy of the policy issuance process.

## Project Overview
### Objective
The primary objective of this project was to create an API that automates the process of validating, extracting, and storing data from insurance quote Excel files into a database, thereby eliminating manual data entry and speeding up the policy issuance process.

### Current Process
1. Underwriter (UW) provides the salesperson with a quotation Excel template.
2. The salesperson fills in customer and risk details in the Excel sheet.
3. The filled Excel sheet is sent back to the UW for quote calculation and approval.
4. The UW calculates the quote and sends it back to the salesperson.
5. The salesperson informs the customer of the calculated quote and collects a cheque.
6. The salesperson submits the cheque along with the quote sheet to the operations team.
7. The operations team manually enters data from the Excel sheet into the front-end application.
8. The operations team issues the policy.

### Challenges with Current Process
i. Inconsistent Excel quote templates due to:
ii. Changes in formulas by salespersons.
iii. Use of obsolete templates.
iv. Time-consuming manual data entry.
v. Error-prone manual data entry.

### Modified Process
1. Underwriter (UW) provides the salesperson with a quotation Excel template.
2. The salesperson fills in customer and risk details in the Excel sheet.
3. The filled Excel sheet is sent back to the UW for quote calculation and approval.
4. The UW calculates the quote and sends it back to the salesperson.
5. The salesperson informs the customer of the calculated quote and collects a cheque.
6. The salesperson submits the cheque along with the quote sheet to the operations team.
7. The application calls the API, which validates the quotation template, extracts the data, and populates the database.
8. The operations team issues the policy.

Advantages of Modified Process
4 Factor Validation:
File format check (.xlsx).
Unique hidden encrypted value check in a random cell.
Template format check using 5 constant labels.
Sheet properties check (creator and title).
Automated Data Entry:
Saves time (from 2 hours to 2 minutes per proposal).
Eliminates data entry errors.
Managing Redundancy in Lead ID
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
