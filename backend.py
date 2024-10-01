from flask import Flask, render_template, request, redirect
import pandas as pd
from openpyxl import load_workbook
import os

app = Flask(__name__)

FILE_PATH = 'responses.xlsx'

# Create an Excel file if it doesn't exist
def create_excel_file():
    if not os.path.exists(FILE_PATH):
        df = pd.DataFrame(columns=['Name', 'Email', 'Message'])
        df.to_excel(FILE_PATH, index=False)

# Save the form data to the Excel file
def save_to_excel(name, email, message):
    create_excel_file()  # Ensure the file exists
    wb = load_workbook(FILE_PATH)
    sheet = wb.active
    sheet.append([name, email, message])
    wb.save(FILE_PATH)

# Home route (form page)
@app.route('/')
def form():
    return render_template('form.html')

# Form submission route
@app.route('/submit', methods=['POST'])
def submit():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        message = request.form['message']

        # Save data to Excel
        save_to_excel(name, email, message)

        return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)
