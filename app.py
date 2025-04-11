from flask import Flask, render_template, request, redirect, url_for, send_file
import os
from result import process_roll_numbers
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        course = request.form['course']
        sem = int(request.form['semester'])

        csv_file = request.files['csv_file']
        if not csv_file or not csv_file.filename.endswith('.csv'):
            return "Please upload a valid CSV file", 400
        file_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
        csv_file.save(file_path)

        try:
            output_path = process_roll_numbers(file_path, course, sem, OUTPUT_FOLDER)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Error processing file: {str(e)}", 500

    return render_template('index.html')

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Default to 5000 if PORT not set
    app.run(host="0.0.0.0", port=port, debug=False)  # Disable debug in production