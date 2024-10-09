from flask import Flask, render_template, request, redirect
import pandas as pd

app = Flask(__name__)

# Path to store the form data in an Excel file
EXCEL_FILE = 'form_data.xlsx'

# Dictionary to hold WhatsApp group links for each course
whatsapp_links = {
    "Python": "https://chat.whatsapp.com/L6pQR5rPQHN7Hkmimxbjrw",
    "Excel": "https://chat.whatsapp.com/L6pQR5rPQHN7Hkmimxbjrw",
    "PowerBI": "https://chat.whatsapp.com/L6pQR5rPQHN7Hkmimxbjrw",
    "Tableau": "https://chat.whatsapp.com/L6pQR5rPQHN7Hkmimxbjrw",
    "Other": "https://chat.whatsapp.com/L6pQR5rPQHN7Hkmimxbjrw"
}

# Initialize Excel file with headers if it doesn't exist
def initialize_excel():
    try:
        pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Name', 'Email', 'Contact', 'Address', 'Course'])
        df.to_excel(EXCEL_FILE, index=False)

# Route for the form page
@app.route('/')
def form():
    return render_template('form.html')

# Route to handle form submission
@app.route('/submit', methods=['POST'])
def submit_form():
    name = request.form['name']
    email = request.form['email']
    contact = request.form['contact']
    address = request.form['address']
    course = request.form['course']

    # Save the data to Excel
    df = pd.read_excel(EXCEL_FILE)
    new_data = pd.DataFrame([{'Name': name, 'Email': email, 'Contact': contact, 'Address': address, 'Course': course}])
    df = pd.concat([df, new_data], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    # Get corresponding WhatsApp link
    whatsapp_link = whatsapp_links.get(course, whatsapp_links["Other"])
    return redirect(whatsapp_link)

if __name__ == '__main__':
    initialize_excel()
    app.run(debug=True)
