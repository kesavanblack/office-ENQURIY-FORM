from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)

# File paths for Excel storage
PRIMARY_EXCEL_FILE = "primary_storage.xlsx"
TEMP_DOWNLOAD_FILE = "download_storage.xlsx"

# Ensure primary storage exists
if not os.path.exists(PRIMARY_EXCEL_FILE):
    # Create an empty DataFrame with predefined columns
    columns = ["Name", "Age", "Sex", "DOB", "Email", "Mobile", "Address", "College Name", 
               "Degree", "Year of Pass Out", "Technical Skills", "Areas of Interest", 
               "Category of Position", "Requirements"]
    df = pd.DataFrame(columns=columns)
    df.to_excel(PRIMARY_EXCEL_FILE, index=False)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Collect form data
    form_data = {
        "Name": request.form['name'],
        "Age": request.form['age'],
        "Sex": request.form['sex'],
        "DOB": request.form['dob'],
        "Email": request.form['email'],
        "Mobile": request.form['mobile'],
        "Address": request.form['address'],
        "College Name": request.form['college_name'],
        "Degree": request.form['degree'],
        "Year of Pass Out": request.form['year_of_passout'],
        "Technical Skills": request.form['technical_skills'],
        "Areas of Interest": request.form['areas_of_interest'],
        "Category of Position": request.form['category_of_position'],
        "Requirements": request.form['requirements']
    }

    # Validate data (example for 'Age' and 'Email')
    if not form_data['Age'].isdigit() or int(form_data['Age']) <= 0:
        return "Invalid Age", 400
    if "@" not in form_data['Email']:
        return "Invalid Email", 400

    # Load existing data and append new record
    df = pd.read_excel(PRIMARY_EXCEL_FILE)
    df = pd.concat([df, pd.DataFrame([form_data])], ignore_index=True)

    # Save updated data back to primary storage
    df.to_excel(PRIMARY_EXCEL_FILE, index=False)

    return redirect(url_for('index'))

@app.route('/export', methods=['GET'])
def export():
    # Read the primary storage file
    df = pd.read_excel(PRIMARY_EXCEL_FILE)

    # Create a BytesIO object to hold the Excel file
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)  # Rewind the buffer to the beginning

    # Send the Excel file to the user
    return send_file(output, download_name="enquiry_data.xlsx", as_attachment=True)

@app.route('/view', methods=['GET'])
def view():
    # Load primary storage file
    df = pd.read_excel(PRIMARY_EXCEL_FILE)

    # Convert DataFrame to HTML for rendering
    table_html = df.to_html(classes='table table-striped', index=False)
    return f"<h1>Stored Enquiries</h1>{table_html}<br><a href='/'>Back to Form</a>"

if __name__ == '__main__':
    app.run(debug=True)
