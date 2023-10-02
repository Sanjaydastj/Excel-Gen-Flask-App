from flask import Flask, request, jsonify, render_template
import os
import json
import tempfile
import win32com.client as win32
import pythoncom  # Import the pythoncom module

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_excel', methods=['POST'])
def generate_excel():
    # Initialize COM
    pythoncom.CoInitialize()

    # Check if the 'json_file' field is in the request
    if 'json_file' not in request.files:
        return jsonify({'error': 'No JSON file provided'})

    uploaded_file = request.files['json_file']

    # Check if the user uploaded a file with a JSON extension
    if uploaded_file.filename == '' or not uploaded_file.filename.endswith('.json'):
        return jsonify({'error': 'Invalid or missing JSON file'})

    # Step 1.1: Read the uploaded JSON file
    try:
        json_data = json.loads(uploaded_file.read().decode('utf-8'))
    except Exception as e:
        return jsonify({'error': f'Error reading JSON file: {str(e)}'})

    # Step 1.2: Flatten the data
    rows = []

    if 'attributes' in json_data:
        attributes_list = json_data['attributes']

        for record in attributes_list:
            name = record.get('beAttrLabel', '')  # Use get() to provide a default value if the key is missing
            dataType = record.get('dataType', '')
            if dataType == 'String':
                dataType = '1'    
            else:
                dataType = "" 
            isSmartAttribute = record.get('isSmartAttribute', '')  # Provide a default value if needed
            displayType = '1'
            length = record.get('length', '')  # Provide a default value if needed
            
            isMandatory = record.get('isMandatory', '')  # Provide a default value if needed
            rows.append([name,dataType, length,displayType, isSmartAttribute, isMandatory])

    # Step 2: Insert records into an Excel Spreadsheet
    ExcelApp = win32.Dispatch('Excel.Application')
    ExcelApp.Visible = True

    wb = ExcelApp.Workbooks.Add()
    ws = wb.Worksheets(1)

    header_labels = ('beAttrLabel', 'length', 'isSmartAttribute', 'isMandatory')

    # Insert header labels
    for indx, val in enumerate(header_labels):
        ws.Cells(1, indx + 1).Value = val

    # Insert records
    row_tracker = 2
    column_size = len(header_labels)

    for row in rows:
        ws.Range(
            ws.Cells(row_tracker, 1),
            ws.Cells(row_tracker, column_size)
        ).value = row
        row_tracker += 1

    # Save Excel file to a temporary directory
    # with tempfile.TemporaryDirectory() as tmp_dir:
    excel_output_path = os.path.join(r'C:\Users\hp\Desktop\Excel Gen App\Output', 'Json_output.xlsx')
    wb.SaveAs(excel_output_path, 51)
    wb.Close()

        # Return the path to the generated Excel file
    return jsonify({'message': 'Excel generated successfully', 'file_path': excel_output_path})

if __name__ == '__main__':
    app.run(debug=True)