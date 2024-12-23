# Quote Extraction Automation API

This repository contains the code and documentation for the Quote Extraction Automation API developed during my internship at HDFC ERGO from April 2022 to June 2022. The project aimed to automate the validation, extraction, and storage of data from insurance quote Excel templates, improving the efficiency and accuracy of the policy issuance process.

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Challenges with Current Process](#challenges-with-current-process)
3. [Modified Process](#modified-process)
4. [Advantages of Modified Process](#advantages-of-modified-process)
5. [Products](#products)
6. [API Code Flow](#api-code-flow)
7. [Demo](#demo)
8. [Software Specifications](#software-specifications)

---

## Project Overview

### Objective
The primary objective of this project was to create an API that automates the process of validating, extracting, and storing data from insurance quote Excel files into a database, thereby eliminating manual data entry and speeding up the policy issuance process.

### Current Process
![Current Process](https://github.com/KunalSachdev2005/Quote_Extraction_Automation_API/blob/main/media/Current_Process.png)

---

## Challenges with Current Process

1. **Inconsistent Excel quote templates due to:**
   - Changes in formulas by salespersons.
   - Use of obsolete templates.
2. **Time-consuming & error-prone manual data entry.**
3. **Redundancy in lead ID:**
   - Each customer is assigned a unique lead ID.
   - A customer might upload the same Excel file multiple times, causing redundancy.

---

## Modified Process

![Modified Process](https://github.com/KunalSachdev2005/Quote_Extraction_Automation_API/blob/main/media/Modified_Process.png)

---

## Advantages of Modified Process

1. **4 Factor Validation:**
   - File format check (.xlsx).
   - Unique hidden encrypted value check in a random cell.
   - Template format check using 5 constant labels.
   - Sheet properties check (creator and title).

2. **Automated Data Entry:**
   - Saves time (from 2 hours to 2 minutes per proposal).
   - Eliminates data entry errors.

3. **Managing Redundancy in Lead ID:**
   - Introduced an "isactive" column in each database table.
   - Ensures the most recent activity is marked as active.

---

## Products

The API handles the following insurance products:

1. **Standard Fire & Special Perils Policy (SFSP)**
2. **Burglary Insurance**
3. **Plate Glass Insurance**

For each product, the API processes sections such as:

- Customer Details
- Main Policies
- Add-On Covers
- Terms & Conditions
- Deductibles
- Warranties
- Clauses & Conditions
- Exclusions & Subjectivities
- Supplementary Clauses & Conditions

---

## API Code Flow

![API Code Flow](https://github.com/KunalSachdev2005/Quote_Extraction_Automation_API/blob/main/media/api_code_flow.png)

---

## Demo

![Demo](https://github.com/KunalSachdev2005/Quote_Extraction_Automation_API/blob/main/media/demo.png)

---

## Software Specifications

- **Python**: 3.8.8
- **Database**: Oracle 19c
- **Libraries:**
  - Flask: 2.1.2
  - Pandas: 1.4.2
  - NumPy: 1.22.4
  - os: 0.6.3
  - sqlalchemy: 1.3.24
  - openpyxl: 3.0.9
  - Werkzeug: 2.1.2
  - cx_Oracle: 8.3.0
    
