<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Quote Extraction Automation API</title>
</head>
<body>
    <h1>Quote Extraction Automation API</h1>
    <p>This repository contains the code and documentation for the Quote Extraction Automation API developed during my internship at HDFC ERGO from April 2022 to June 2022. The project aimed to automate the validation, extraction, and storage of data from insurance quote Excel templates, improving the efficiency and accuracy of the policy issuance process.</p>

    <h2>Project Overview</h2>

    <h3>Objective</h3>
    <p>The primary objective of this project was to create an API that automates the process of validating, extracting, and storing data from insurance quote Excel files into a database, thereby eliminating manual data entry and speeding up the policy issuance process.</p>

    <h3>Current Process</h3>
    <ol>
        <li>Underwriter (UW) provides the salesperson with a quotation Excel template.</li>
        <li>The salesperson fills in customer and risk details in the Excel sheet.</li>
        <li>The filled Excel sheet is sent back to the UW for quote calculation and approval.</li>
        <li>The UW calculates the quote and sends it back to the salesperson.</li>
        <li>The salesperson informs the customer of the calculated quote and collects a cheque.</li>
        <li>The salesperson submits the cheque along with the quote sheet to the operations team.</li>
        <li>The operations team manually enters data from the Excel sheet into the front-end application.</li>
        <li>The operations team issues the policy.</li>
    </ol>

    <h3>Challenges with Current Process</h3>
    <ul>
        <li>Inconsistent Excel quote templates due to:
            <ul>
                <li>Changes in formulas by salespersons.</li>
                <li>Use of obsolete templates.</li>
            </ul>
        </li>
        <li>Time-consuming manual data entry.</li>
        <li>Error-prone manual data entry.</li>
    </ul>

    <h3>Modified Process</h3>
    <ol>
        <li>Underwriter provides the salesperson with a quotation Excel template.</li>
        <li>The salesperson fills in customer and risk details in the Excel sheet.</li>
        <li>The filled Excel sheet is sent back to the UW for quote calculation and approval.</li>
        <li>The UW calculates the quote and sends it back to the salesperson.</li>
        <li>The salesperson informs the customer of the calculated quote and collects a cheque.</li>
        <li>The salesperson submits the cheque along with the quote sheet to the application.</li>
        <li>The application calls the API, which validates the quotation template, extracts the data, and populates the database.</li>
        <li>The operations team issues the policy.</li>
    </ol>

    <h3>Advantages of Modified Process</h3>
    <ul>
        <li><strong>4 Factor Validation:</strong>
            <ul>
                <li>File format check (.xlsx).</li>
                <li>Unique hidden encrypted value check in a random cell.</li>
                <li>Template format check using 5 constant labels.</li>
                <li>Sheet properties check (creator and title).</li>
            </ul>
        </li>
        <li><strong>Automated Data Entry:</strong>
            <ul>
                <li>Saves time (from 2 hours to 2 minutes per proposal).</li>
                <li>Eliminates data entry errors.</li>
            </ul>
        </li>
    </ul>

    <h3>Managing Redundancy in Lead ID</h3>
    <p>To handle redundancy in lead IDs:</p>
    <ul>
        <li>An "isactive" column is introduced in each database table.</li>
        <li>Upon data upload, the API checks for existing lead IDs and updates the "isactive" status accordingly, ensuring the most recent activity is marked as active.</li>
    </ul>

    <h3>Products</h3>
    <p>The API handles the following insurance products:</p>
    <ol>
        <li>Standard Fire & Special Perils Policy (SFSP)</li>
        <li>Burglary Insurance</li>
        <li>Plate Glass Insurance</li>
    </ol>
    <p>For each product, the API processes sections such as Customer Details, Main Policies, Add-On Covers, Terms & Conditions, Deductibles, Warranties, Clauses & Conditions, Exclusions & Subjectivities, and Supplementary Clauses & Conditions.</p>

    <h2>API Code Flow</h2>
    <ol>
        <li><strong>File Upload:</strong> The user uploads the quote file.</li>
        <li><strong>Quote Format Validation:</strong> The API validates the file format and constant sheet values.</li>
        <li><strong>Quote Extraction:</strong> The API extracts data and uploads it to the database.</li>
    </ol>

    <h2>Demo</h2>
    <p>The repository includes a demo showcasing the front-end for file upload and the resulting output message.</p>

    <h2>Software Specifications</h2>
    <ul>
        <li><strong>Python:</strong> 3.8.8</li>
        <li><strong>Database:</strong> Oracle 19c</li>
        <li><strong>Libraries:</strong>
            <ul>
                <li>Flask: 2.1.2</li>
                <li>Pandas: 1.4.2</li>
                <li>NumPy: 1.22.4</li>
                <li>os: 0.6.3</li>
                <li>sqlalchemy: 1.3.24</li>
                <li>openpyxl: 3.0.9</li>
                <li>Werkzeug: 2.1.2</li>
                <li>cx_Oracle: 8.3.0</li>
            </ul>
        </li>
    </ul>

    <h2>Acknowledgements</h2>
    <p>Special thanks to:</p>
    <ul>
        <li><strong>Mr. Naresh Jha</strong> (Senior VP-IT, HDFC ERGO) for the opportunity.</li>
        <li><strong>Mr. Ketan Khandagale</strong> and <strong>Mr. Swapnil Gunjal</strong> for their mentorship and support.</li>
        <li>My parents for their continuous encouragement.</li>
    </ul>
</body>
</html>
