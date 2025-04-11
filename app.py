from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import pandas as pd
from result import process_roll_numbers

app = Flask(__name__)
UPLOAD_FOLDER = 'Uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        input_method = request.form['input_method']
        course = request.form['course']
        sem = int(request.form['semester'])

        if input_method == 'upload':
            csv_file = request.files['csv_file']
            if not csv_file or not csv_file.filename.endswith('.csv'):
                return "Please upload a valid CSV file", 400
            file_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)
            csv_file.save(file_path)
        else:  # input_method == 'generate'
            college_code = request.form['college_code']
            branch = request.form['branch']
            year = request.form['year'].zfill(2)  # Ensure two digits (e.g., '23')
            student_count = int(request.form['student_count'])

            # Validate inputs
            if student_count < 1 or student_count > 500:
                return "Student count must be between 1 and 500", 400
            if not year.isdigit() or not (0 <= int(year) <= 99):
                return "Year must be two digits (e.g., 23 for 2023)", 400

            # Generate roll numbers
            roll_numbers = [
                f"{college_code}{branch}{year}{str(i).zfill(3)}"
                for i in range(1001, 1001 + student_count)
            ]
            # Save to temporary CSV
            file_path = os.path.join(UPLOAD_FOLDER, 'generated_roll_numbers.csv')
            pd.DataFrame(roll_numbers, columns=['Roll Number']).to_csv(file_path, index=False)

        try:
            output_path = process_roll_numbers(file_path, course, sem, OUTPUT_FOLDER)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Error processing file: {str(e)}", 500

    return render_template('index.html')

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)