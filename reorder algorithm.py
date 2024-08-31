from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os

app = Flask(__name__)

# Ensure upload folder exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def process_file(filepath):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(filepath, sheet_name='Reorder Report Branch 2520', skiprows=2)
    
    # Add  custom reorder rules here
    # Example: Filter items that need reordering
    reorder_df = df[df['Req'] > 0]  # Modify this condition as per rules
    
    # Save the results to a new Excel file
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reorder_list.xlsx')
    reorder_df.to_excel(output_path, index=False)

    return output_path

# HTML Template
UPLOAD_PAGE_HTML = '''
<!DOCTYPE html>
<html>
<head>
    <title>Upload File for Reorder Processing</title>
</head>
<body>
    <h2>Upload Your Reorder Report</h2>
    <form method="post" action="/upload" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <input type="submit" value="Upload and Process">
    </form>
</body>
</html>
'''

@app.route('/')
def upload_file():
    return render_template_string(UPLOAD_PAGE_HTML)

@app.route('/upload', methods=['POST'])
def handle_file_upload():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']
