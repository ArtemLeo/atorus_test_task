from flask import Flask, render_template, request, send_file
import pandas as pd
import pyreadstat
import chardet
import os
import subprocess

app = Flask(__name__)
TEMP_FOLDER = 'temp/'  # Temporary folder for saving files
os.makedirs(TEMP_FOLDER, exist_ok=True)

# Global variables to store data and file type
global_data = None
file_type_global = None

# Function to convert CPORT to XPT
def convert_cport_to_xpt(input_path, output_path):
    """
    Викликає зовнішній інструмент для конвертації CPORT у XPT.
    """
    try:
        subprocess.run(
            ["stattransfer", f"/in={input_path}", f"/out={output_path}", "/cport-to-xpt"],
            check=True
        )
    except subprocess.CalledProcessError as e:
        raise ValueError(f"Помилка під час конвертації CPORT: {e}")

# Separate functions for each file format
def read_csv(file):
    raw_data = file.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    file.seek(0)
    return pd.read_csv(file, encoding=encoding, delimiter='$')

def read_excel(file):
    return pd.read_excel(file)

def read_sas7bdat(file):
    temp_file_path = os.path.join(TEMP_FOLDER, file.filename)
    file.save(temp_file_path)  # Save the uploaded file temporarily
    data, _ = pyreadstat.read_sas7bdat(temp_file_path)
    os.remove(temp_file_path)  # Delete the temporary file after reading
    return data

def read_xpt(file):
    temp_file_path = os.path.join(TEMP_FOLDER, file.filename)
    file.save(temp_file_path)  # Save the uploaded file temporarily

    try:
        # Try to read as standard XPT
        data, _ = pyreadstat.read_xport(temp_file_path)
    except Exception as e:
        if "CPORT" in str(e):  # If it's a CPORT file
            converted_file_path = temp_file_path + ".xpt"
            convert_cport_to_xpt(temp_file_path, converted_file_path)  # Convert to XPT
            try:
                data, _ = pyreadstat.read_xport(converted_file_path)
            except Exception as e:
                os.remove(temp_file_path)
                os.remove(converted_file_path)
                raise ValueError(f"Не вдалося прочитати сконвертований файл: {e}")
            os.remove(converted_file_path)  # Remove converted file
        else:
            os.remove(temp_file_path)
            raise ValueError(f"Помилка читання XPT: {e}")

    os.remove(temp_file_path)  # Remove original file after processing
    return data

# Function to clean the data
def clean_table(data):
    data = data.dropna(axis=1, how='all')  # Drop columns where all values are NaN
    data = data.fillna('N/A')  # Fill NaN values with 'N/A'
    data.columns = [col.strip().replace(' ', '_').lower() for col in data.columns]  # Rename columns to snake_case
    for col in data.select_dtypes(include=['object']).columns:  # Remove unwanted characters
        data[col] = data[col].str.replace(r'[\n\r]+', ' ', regex=True).str.strip()
        data[col] = data[col].str.replace(r'\s+', ' ', regex=True)  # Normalize spaces
    return data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/view', methods=['POST'])
def view_table():
    global global_data, file_type_global
    try:
        file = request.files['file']
        file_type = request.form['file_type']
        if not file:
            return "No file uploaded!", 400

        if file_type == 'xpt':
            try:
                data = read_xpt(file)
            except ValueError as e:
                if "CPORT" in str(e):
                    return (
                        "Файл є форматом CPORT, який не підтримується. "
                        "Будь ласка, конвертуйте його у XPT або SAS7BDAT перед завантаженням.",
                        400
                    )
                else:
                    return f"Помилка обробки XPT файлу: {e}", 400
        elif file_type == 'csv':
            data = read_csv(file)
        elif file_type == 'excel':
            data = read_excel(file)
        elif file_type == 'sas7bdat':
            data = read_sas7bdat(file)
        else:
            return "Unsupported file type selected!", 400

        data = clean_table(data)  # Clean the data
        html_table = data.to_html(classes='data', index=False, escape=False)
        global_data = data  # Store cleaned data globally
        file_type_global = file_type
        return render_template('table.html', table_html=html_table)
    except Exception as e:
        return f"Error processing file: {str(e)}", 500


@app.route('/save')
def save_file():
    global global_data
    if global_data is None:
        return "No data to save!", 400

    save_path = os.path.join(TEMP_FOLDER, 'saved_file.xlsx')  # Always save as Excel
    try:
        global_data.to_excel(save_path, index=False, engine='openpyxl')  # Save as Excel
        return send_file(save_path, as_attachment=True)
    except Exception as e:
        return f"Error saving file: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
