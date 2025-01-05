from flask import Flask, render_template, request, send_file, session
import pandas as pd
import pyreadstat
import chardet
import os
import subprocess
import csv
import uuid

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'

# Папка для зберігання сконвертованих (готових до вивантаження) файлів
CONVERTED_FOLDER = 'converted_files/'
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# ------------------------------------------------------------------------------
# ФУНКЦІЇ ДЛЯ ОБРОБКИ РІЗНИХ ФОРМАТІВ
# ------------------------------------------------------------------------------

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

def read_csv(file):
    """
    Гнучке читання CSV:
      1) Автоматично визначаємо роздільник через csv.Sniffer (fallback — '$').
      2) Визначаємо кодування через chardet (fallback — 'utf-8').
    """
    raw_data = file.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding'] or 'utf-8'

    file.seek(0)
    snippet = file.read(2048)
    file.seek(0)

    try:
        snippet_str = snippet.decode(encoding, errors='replace')
    except:
        snippet_str = snippet.decode('utf-8', errors='replace')

    possible_delimiters = ['$',';',',','\t','|',':']
    try:
        dialect = csv.Sniffer().sniff(snippet_str, delimiters=possible_delimiters)
        delimiter = dialect.delimiter
    except Exception:
        delimiter = '$'

    file.seek(0)
    return pd.read_csv(file, encoding=encoding, delimiter=delimiter)

def read_excel(file):
    """
    Читання Excel (XLS, XLSX) за допомогою pandas.
    """
    return pd.read_excel(file)

def read_sas7bdat(file):
    """
    Pyreadstat вимагає файл на диску.
    Якщо це Flask FileStorage (має .filename і .save()), викликаємо .save().
    Інакше зберігаємо вручну.
    """
    if hasattr(file, 'save') and hasattr(file, 'filename'):
        temp_file_path = os.path.join(CONVERTED_FOLDER, file.filename)
        file.save(temp_file_path)
    else:
        random_name = f"test_sas7bdat_{uuid.uuid4().hex}.sas7bdat"
        temp_file_path = os.path.join(CONVERTED_FOLDER, random_name)
        with open(temp_file_path, "wb") as out_f:
            out_f.write(file.read())

    data, _ = pyreadstat.read_sas7bdat(temp_file_path)
    os.remove(temp_file_path)
    return data

def read_xpt(file):
    """
    Може бути стандартний XPT або CPORT.
    Якщо це CPORT (у тексті помилки “CPORT”), тоді конвертуємо у XPT.

    Аналогічно read_sas7bdat: перевіряємо, чи є file.filename/file.save().
    Якщо ні — зберігаємо вручну в random_file.xpt
    """
    if hasattr(file, 'save') and hasattr(file, 'filename'):
        temp_file_path = os.path.join(CONVERTED_FOLDER, file.filename)
        file.save(temp_file_path)
    else:
        random_name = f"test_xpt_{uuid.uuid4().hex}.xpt"
        temp_file_path = os.path.join(CONVERTED_FOLDER, random_name)
        with open(temp_file_path, "wb") as out_f:
            out_f.write(file.read())

    try:
        data, _ = pyreadstat.read_xport(temp_file_path)
    except Exception as e:
        # Якщо це CPORT, у повідомленні часто є "CPORT"
        if "CPORT" in str(e):
            converted_file_path = temp_file_path + ".xpt"
            convert_cport_to_xpt(temp_file_path, converted_file_path)
            try:
                data, _ = pyreadstat.read_xport(converted_file_path)
            except Exception as e2:
                # Видаляємо обидва файли, якщо вони існують
                if os.path.exists(temp_file_path):
                    os.remove(temp_file_path)
                if os.path.exists(converted_file_path):
                    os.remove(converted_file_path)
                raise ValueError(f"Не вдалося прочитати сконвертований файл: {e2}")

            # Видаляємо конвертований файл, якщо він існує
            if os.path.exists(converted_file_path):
                os.remove(converted_file_path)
        else:
            # Якщо помилка не схожа на CPORT, просто видаляємо temp_file_path
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            raise ValueError(f"Помилка читання XPT: {e}")

    if os.path.exists(temp_file_path):
        os.remove(temp_file_path)

    return data

def clean_table(data):
    """
    1) Видаляє порожні колонки
    2) Заповнює NaN як 'N/A'
    3) Приводить назви колонок до snake_case
    4) Тримає під контролем зайві пробіли та перенос рядків
    """
    data = data.dropna(axis=1, how='all')
    data = data.fillna('N/A')
    data.columns = [col.strip().replace(' ', '_').lower() for col in data.columns]

    for col in data.select_dtypes(include=['object']).columns:
        data[col] = data[col].str.replace(r'[\n\r]+', ' ', regex=True).str.strip()
        data[col] = data[col].str.replace(r'\s+', ' ', regex=True)

    return data

# ------------------------------------------------------------------------------
# МАПІНГ РОЗШИРЕНЬ НА ФУНКЦІЇ
# ------------------------------------------------------------------------------

EXTENSION_TO_READER = {
    '.csv': read_csv,
    '.xpt': read_xpt,
    '.sas7bdat': read_sas7bdat,
    '.xlsx': read_excel,
    '.xls': read_excel
}

# ------------------------------------------------------------------------------
# МАРШРУТИ
# ------------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/view', methods=['POST'])
def view_table():
    try:
        file = request.files['file']
        if not file:
            return "No file uploaded!", 400

        original_filename = file.filename
        ext = os.path.splitext(original_filename)[1].lower()
        reader_func = EXTENSION_TO_READER.get(ext)
        if not reader_func:
            return "Unsupported file type selected!", 400

        data = reader_func(file)
        data = clean_table(data)

        base_name, _ = os.path.splitext(original_filename)
        excel_filename = base_name + ".xlsx"
        excel_path = os.path.join(CONVERTED_FOLDER, excel_filename)
        data.to_excel(excel_path, index=False, engine='openpyxl')

        html_table = data.to_html(classes='data', index=False, escape=False)
        session['excel_filename'] = excel_filename

        return render_template('table.html', table_html=html_table)

    except ValueError as e:
        return str(e), 400
    except Exception as e:
        return f"Error processing file: {str(e)}", 500

@app.route('/save')
def save_file():
    excel_filename = session.get('excel_filename')
    if not excel_filename:
        return "No data to save!", 400

    excel_path = os.path.join(CONVERTED_FOLDER, excel_filename)
    if not os.path.exists(excel_path):
        return f"File {excel_filename} not found on server!", 404

    return send_file(excel_path, as_attachment=True, download_name=excel_filename)

if __name__ == '__main__':
    app.run(debug=True)
