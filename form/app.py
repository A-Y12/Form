from flask import Flask, request, render_template_string
import pandas as pd
import os

app = Flask(__name__)
excel_file = 'Warehouse materiala issue & return Details.xlsx'

form_html = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Project Details Form</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f9;
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            overflow: auto;
        }
        .container {
            background-color: #fff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            max-width: 800px;
            width: 100%;
        }
        h1 {
            text-align: center;
            color: #333;
            margin-bottom: 20px;
        }
        form {
            display: flex;
            flex-direction: column;
        }
        label {
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
        }
        input, select {
            margin-bottom: 15px;
            padding: 10px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 16px;
            width: calc(100% - 22px);
        }
        input[type="submit"] {
            background-color: #007bff;
            color: #fff;
            border: none;
            cursor: pointer;
            transition: background-color 0.3s;
            padding: 15px;
            font-size: 18px;
        }
        input[type="submit"]:hover {
            background-color: #0056b3;
        }
        .row {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
        }
        .row > div {
            flex: 0 0 48%;
            box-sizing: border-box;
        }
        .row > div input, .row > div select {
            width: 100%;
        }
    </style>
    <script>
        function validateForm() {
            let year = document.getElementById('year').value;
            if (year < 2000 || year > 2100) {
                alert("Year must be between 2000 and 2100");
                return false;
            }
            return true;
        }
        function showError(message) {
            alert(message);
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>Project Details Form</h1>
        <form action="/submit" method="post" onsubmit="return validateForm()">
            <div class="row">
                <div>
                    <label for="project_name">Project Name:</label>
                    <input type="text" id="project_name" name="project_name" required>
                </div>
                <div>
                    <label for="speed_code">Speed Code:</label>
                    <input type="text" id="speed_code" name="speed_code" required>
                </div>
            </div>
            <div class="row">
                <div>
                    <label for="project_start_date">Project Start Date:</label>
                    <input type="date" id="project_start_date" name="project_start_date" required>
                </div>
                <div>
                    <label for="project_end_date">Project End Date:</label>
                    <input type="date" id="project_end_date" name="project_end_date" required>
                </div>
            </div>
            <div class="row">
                <div>
                    <label for="pr_no">Pr. No:</label>
                    <input type="text" id="pr_no" name="pr_no" required>
                </div>
                <div>
                    <label for="po_no">PO No:</label>
                    <input type="text" id="po_no" name="po_no" required>
                </div>
            </div>
            <div class="row">
                <div>
                    <label for="year">Year:</label>
                    <input type="number" id="year" name="year" required>
                </div>
                <div>
                    <label for="date">Date:</label>
                    <input type="date" id="date" name="date" required>
                </div>
            </div>
            <div class="row">
                <div>
                    <label for="document">Document:</label>
                    <input type="text" id="document" name="document" required>
                </div>
                <div>
                    <label for="open_close">Open/Close:</label>
                    <select id="open_close" name="open_close" required>
                        <option value="Open">Open</option>
                        <option value="Close">Close</option>
                    </select>
                </div>
            </div>
            <input type="submit" value="Submit">
        </form>
        <script>
            {% if error_message %}
            showError("{{ error_message }}");
            {% endif %}
        </script>
    </div>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(form_html)

@app.route('/submit', methods=['POST'])
def submit():
    try:
        project_name = request.form['project_name']
        speed_code = request.form['speed_code']
        project_start_date = request.form['project_start_date']
        project_end_date = request.form['project_end_date']
        pr_no = request.form['pr_no']
        po_no = request.form['po_no']
        year = request.form['year']
        date = request.form['date']
        document = request.form['document']
        open_close = request.form['open_close']

        data = {
            'Project Name': project_name,
            'Speed Code': speed_code,
            'Project Start Date': project_start_date,
            'Project End Date': project_end_date,
            'Pr. No': pr_no,
            'PO No': po_no,
            'Year': year,
            'Date': date,
            'Document': document,
            'Open/Close': open_close
        }

        # Check if the file exists
        if os.path.exists(excel_file):
            try:
                # Try to read the existing Excel file
                df = pd.read_excel(excel_file)
            except PermissionError:
                return render_template_string(form_html, error_message="Permission denied: Please close the Excel file and try again.")
        else:
            # If the file does not exist, create a new DataFrame
            df = pd.DataFrame(columns=['S.No', 'Project Name', 'Speed Code', 'Project Start Date', 'Project End Date', 'Pr. No', 'PO No', 'Year', 'Date', 'Document', 'Open/Close'])

        # Determine the next serial number
        next_s_no = len(df) + 1

        # Add the serial number to the data
        data['S.No'] = next_s_no

        # Append the new data
        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)

        try:
            # Try to save the updated DataFrame to the Excel file
            df.to_excel(excel_file, index=False)
        except PermissionError:
            return render_template_string(form_html, error_message="Permission denied: Please close the Excel file and try again.")

        return render_template_string(form_html, success_message="Data submitted successfully.")
    except Exception as e:
        return render_template_string(form_html, error_message=f"An error occurred: {e}")

if __name__ == '__main__':
    app.run(debug=True)
