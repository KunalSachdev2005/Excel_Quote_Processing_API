# ********************************************************************************************************************
# Program Name: Quote_API.py
# Description: This progarm is an API to validate, exctract, and store information from BSC Quote Excel into Database.
# Created on 03-06-2022 by Kunal Sachdev
# ********************************************************************************************************************


# Importing the required libraries
from timeit import default_timer as timer
import os
from wsgiref.validate import validator
from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
from Extraction import oricon_data_retriever_uploader
from Validation import validation

UPLOAD_FOLDER = 'C:\\Users\\Kunal Sachdev\\Desktop\\ext\\UPLOADED_FILES'

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET','POST'])

# Defining the upload file function
def upload_file():
    if request.method == 'POST':
        start = timer()
        file = request.files['file']
        filename = secure_filename(file.filename)
        print(filename)
        ftype = filename.split('.')[1]
        if ftype != "xlsx":
            raise Exception("Invalid File Type (File is not in .xlsx format)")
        else:
            # Calling the validation function
            if validation(filename):
                try:
                    # Calling the extractor function
                    oricon_data_retriever_uploader(filename)
                except:
                    return ("EXTRACTION FAILED. DATA NOT PUSHED INTO DATABASE")
            else:
                raise Exception("VALIDATION FAILED.  PLEASE UPLOAD A VALID EXCEL FILE.")
            end = timer()
            time_elapsed = end-start
            # Calculating time taken to execute
            time_elapsed = round(time_elapsed, 2)
            return ("DATA PUSHED SUCCESSFULLY \n (TIME TAKEN:" +str(time_elapsed)+ " SECONDS).  THANK YOU FOR USING OUR SERVICE!"),200
    
    # html code
    return '''
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    body {
    margin: 0;
    font-family: "Times New Roman", Times, serif;
    }
    
    .bg-image {
    background-image: url("data:image/jpeg;base64,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");
    background-color: #ffffff;
    height: 750px;
    background-position: left;
    background-repeat: no-repeat;
    background-size: 400px 300px;
    position: relative;
    }
    
    .bg-text {
    text-align: center;
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    color: black;
    }
    
    </style>
    </head>
    <body>
    <div class="bg-image">
    <div class="bg-text">
    <h1 style="font-size:30px">PLEASE UPLOAD QUOTE FILE</h1>
    <form method=post enctype=multipart/form-data>
        <input type=file name=file>
        <input type=submit value=UPLOAD>
    </form>
    <h3></h3>
    </div>
    </div>
    </body>
    '''

	
if __name__ == "__main__":
    app.secret_key = 'super secret key'
    app.run(host='localhost', port=5000, debug=True)