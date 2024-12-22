from flask import Flask, render_template, request, send_file, session
import pandas as pd
import pyreadstat
import chardet
import os
import subprocess
import csv

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
      1) Автоматично намагаємося визначити роздільник через csv.Sniffer
         із заданого списку можливих (поставимо $ першим).
      2) Якщо визначити не вдається, fallback – $.
    """
    # Зчитуємо увесь вміст, визначаємо кодування
    raw_data = file.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']

    # Повертаємо курсор у файлі на початок
    file.seek(0)

    # Зчитуємо невеликий фрагмент (2 кБ) для Sniffer
    snippet = file.read(2048)
    file.seek(0)

    # Декодуємо snippet у рядок
    try:
        snippet_str = snippet.decode(encoding, errors='replace')
    except:
        # Якщо щось пішло не так, fallback – UTF-8
        snippet_str = snippet.decode('utf-8', errors='replace')

    # Набір можливих роздільників (ставимо $ першим)
    possible_delimiters = ['$',';',',','\t','|',':']

    try:
        dialect = csv.Sniffer().sniff(snippet_str, delimiters=possible_delimiters)
        delimiter = dialect.delimiter
    except Exception:
        # Якщо Sniffer не впорався, fallback – $
        delimiter = '$'

    # Знову повертаємось на початок файлу
    file.seek(0)

    # Читаємо CSV з обраним роздільником
    return pd.read_csv(file, encoding=encoding, delimiter=delimiter)

def read_excel(file):
    """
    Читання Excel (XLS або XLSX) із використанням pandas.read_excel.
    Роздільник у цьому форматі не потрібен.
    """
    return pd.read_excel(file)

def read_sas7bdat(file):
    """
    Pyreadstat вимагає, щоб файл був на диску.
    Тому тимчасово зберігаємо його у CONVERTED_FOLDER, а потім видаляємо.
    """
    temp_file_path = os.path.join(CONVERTED_FOLDER, file.filename)
    file.save(temp_file_path)
    data, _ = pyreadstat.read_sas7bdat(temp_file_path)
    os.remove(temp_file_path)
    return data

def read_xpt(file):
    """
    Може бути стандартний XPT або CPORT.
    Якщо це CPORT (у тексті помилки “CPORT”), тоді конвертуємо у XPT.
    """
    temp_file_path = os.path.join(CONVERTED_FOLDER, file.filename)
    file.save(temp_file_path)
    try:
        data, _ = pyreadstat.read_xport(temp_file_path)
    except Exception as e:
        # Якщо це CPORT, у повідомленні часто є слово "CPORT"
        if "CPORT" in str(e):
            converted_file_path = temp_file_path + ".xpt"
            convert_cport_to_xpt(temp_file_path, converted_file_path)
            try:
                data, _ = pyreadstat.read_xport(converted_file_path)
            except Exception as e:
                os.remove(temp_file_path)
                os.remove(converted_file_path)
                raise ValueError(f"Не вдалося прочитати сконвертований файл: {e}")
            os.remove(converted_file_path)
        else:
            os.remove(temp_file_path)
            raise ValueError(f"Помилка читання XPT: {e}")
    os.remove(temp_file_path)
    return data

def clean_table(data):
    """
    Видаляє порожні колонки, заповнює NaN як 'N/A',
    нормалізує рядки (видаляє зайві пробіли та перенос рядків),
    а також робить назви колонок у snake_case.
    """
    data = data.dropna(axis=1, how='all')  # видаляємо повністю порожні стовпці
    data = data.fillna('N/A')              # замінюємо відсутні значення
    data.columns = [col.strip().replace(' ', '_').lower() for col in data.columns]

    for col in data.select_dtypes(include=['object']).columns:
        data[col] = data[col].str.replace(r'[\n\r]+', ' ', regex=True).str.strip()
        data[col] = data[col].str.replace(r'\s+', ' ', regex=True)
    return data

# ------------------------------------------------------------------------------
# МАПІНГ РОЗШИРЕНЬ НА ФУНКЦІЇ
# ------------------------------------------------------------------------------

EXTENSION_TO_READER = {
    '.csv': read_csv,      # CSV із гнучким роздільником
    '.xpt': read_xpt,      # XPT або CPORT (автообробка)
    '.sas7bdat': read_sas7bdat,
    '.xlsx': read_excel,   # Excel (XLSX)
    '.xls': read_excel     # Excel (XLS)
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
        ext = os.path.splitext(original_filename)[1].lower()  # Напр. ".csv"

        # 1. Автоматично обираємо функцію зчитування за розширенням
        reader_func = EXTENSION_TO_READER.get(ext)
        if not reader_func:
            return "Unsupported file type selected!", 400

        # 2. Зчитуємо дані
        data = reader_func(file)

        # 3. Очищаємо дані
        data = clean_table(data)

        # 4. Формуємо назву Excel-файлу з тим самим "коренем"
        base_name, _ = os.path.splitext(original_filename)
        excel_filename = base_name + ".xlsx"
        excel_path = os.path.join(CONVERTED_FOLDER, excel_filename)

        # 5. Зберігаємо у форматі Excel
        data.to_excel(excel_path, index=False, engine='openpyxl')

        # 6. Генеруємо HTML-таблицю для відображення
        html_table = data.to_html(classes='data', index=False, escape=False)

        # 7. Зберігаємо назву Excel-файлу в session
        session['excel_filename'] = excel_filename

        return render_template('table.html', table_html=html_table)

    except ValueError as e:
        # Обробка ситуації з CPORT або інших ValueError
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

    return send_file(
        excel_path,
        as_attachment=True,
        download_name=excel_filename
    )

if __name__ == '__main__':
    app.run(debug=True)
