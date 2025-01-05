# Перевірка та запуск: coverage run -m pytest --disable-warnings -v

import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest
import io
import pandas as pd
from unittest.mock import patch, MagicMock
from app import (
    app,
    read_csv,
    read_excel,
    read_sas7bdat,
    read_xpt,
    clean_table,
    convert_cport_to_xpt,
    EXTENSION_TO_READER,
    CONVERTED_FOLDER
)


# ----------------- ЮНІТ-ТЕСТИ (CSV) ----------------- #

@pytest.fixture
def csv_file_comma():
    return io.BytesIO(b"col1,col2\n1,2\n3,4")

@pytest.fixture
def csv_file_semicolon():
    return io.BytesIO(b"colA;colB\nfoo;bar\nbaz;qux")

@pytest.fixture
def csv_file_dollar():
    return io.BytesIO(b"colX$colY\n100$200\n300$400")

@pytest.fixture
def csv_file_ambiguous():
    """
    Створюємо CSV, де кожен рядок має різні роздільники,
    щоби csv.Sniffer не зміг визначити один єдиний.
    """
    content = b"""col1;col2|col3
val1,val2:val3
abc\tdef:ghi
"""
    return io.BytesIO(content)

def test_read_csv_comma(csv_file_comma):
    df = read_csv(csv_file_comma)
    assert list(df.columns) == ["col1", "col2"]
    assert df.shape == (2, 2)
    assert df.iloc[0, 0] == 1

def test_read_csv_semicolon(csv_file_semicolon):
    df = read_csv(csv_file_semicolon)
    assert list(df.columns) == ["colA", "colB"]
    assert df.shape == (2, 2)
    assert df.iloc[0, 0] == "foo"

def test_read_csv_dollar(csv_file_dollar):
    df = read_csv(csv_file_dollar)
    assert list(df.columns) == ["colX", "colY"]
    assert df.shape == (2, 2)
    assert df.iloc[0, 1] == 200

def test_read_csv_ambiguous(csv_file_ambiguous):
    df = read_csv(csv_file_ambiguous)
    # Якщо Sniffer не визначає delimiter, fallback '$', усе з'єднується в одну колонку
    assert df.shape[1] == 1


# ----------------- ЮНІТ-ТЕСТИ (Excel) ----------------- #

@pytest.fixture
def xlsx_file():
    """
    Створюємо простий Excel-файл у пам'яті
    з колонками A, B та двома рядками.
    """
    df_in = pd.DataFrame({"A": [1, 2], "B": ["foo", "bar"]})
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_in.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    return buffer

def test_read_excel(xlsx_file):
    df_out = read_excel(xlsx_file)
    assert list(df_out.columns) == ["A", "B"]
    assert df_out.shape == (2, 2)
    assert df_out.iloc[0, 1] == "foo"


# ----------------- ЮНІТ-ТЕСТИ (SAS7BDAT) ----------------- #

@patch("app.pyreadstat.read_sas7bdat")
def test_read_sas7bdat(mock_read_sas7bdat, tmp_path):
    """
    Перевіряємо, що read_sas7bdat зберігає файл на диск і викликає pyreadstat.read_sas7bdat().
    """
    mock_df = pd.DataFrame({"col1": [1, 2], "col2": ["a", "b"]})
    mock_meta = MagicMock()
    mock_read_sas7bdat.return_value = (mock_df, mock_meta)

    test_file = tmp_path / "test.sas7bdat"
    test_file.write_text("dummy content")
    with open(test_file, "rb") as f:
        df = read_sas7bdat(f)

    mock_read_sas7bdat.assert_called_once()
    assert df.equals(mock_df)


# ----------------- ЮНІТ-ТЕСТИ (XPT / CPORT) ----------------- #

@patch("app.pyreadstat.read_xport")
def test_read_xpt_normal_xpt(mock_read_xport, tmp_path):
    """
    Якщо файл справжній XPT,
    має прочитатися без конвертації (CPORT).
    """
    mock_df = pd.DataFrame({"x": [10], "y": [20]})
    mock_meta = MagicMock()
    mock_read_xport.return_value = (mock_df, mock_meta)

    test_file = tmp_path / "test.xpt"
    test_file.write_text("dummy xpt content")
    with open(test_file, "rb") as f:
        df = read_xpt(f)

    assert df.equals(mock_df)

@patch("app.convert_cport_to_xpt")
@patch("app.pyreadstat.read_xport")
def test_read_xpt_cport(mock_read_xport, mock_convert_cport_to_xpt, tmp_path):
    """
    Якщо read_xport кидає Exception("Something about CPORT"),
    маємо викликати convert_cport_to_xpt і повторно прочитати файл.
    """
    def side_effect(*args, **kwargs):
        if not hasattr(side_effect, "called"):
            side_effect.called = True
            raise Exception("Something about CPORT")
        else:
            mock_df = pd.DataFrame({"c": [999]})
            mock_meta = MagicMock()
            return mock_df, mock_meta

    mock_read_xport.side_effect = side_effect

    test_file = tmp_path / "test.xpt"
    test_file.write_text("dummy cport content")
    with open(test_file, "rb") as f:
        df = read_xpt(f)

    # Маємо упевнитися, що convert_cport_to_xpt викликано
    mock_convert_cport_to_xpt.assert_called_once()
    assert df.shape == (1, 1)
    assert df.iloc[0, 0] == 999


@patch("subprocess.run")
def test_convert_cport_to_xpt(mock_subprocess_run):
    """
    Перевірка виклику зовнішньої команди для конвертації CPORT -> XPT
    """
    convert_cport_to_xpt("input.cport", "output.xpt")
    mock_subprocess_run.assert_called_once_with(
        ["stattransfer", "/in=input.cport", "/out=output.xpt", "/cport-to-xpt"],
        check=True
    )


# ----------------- ЮНІТ-ТЕСТ clean_table ----------------- #

def test_clean_table():
    df_in = pd.DataFrame({
        "  Col 1 ": [None, " foo "],
        "Col_2": [None, " \n Bar \n"],
        "Empty": [None, None]
    })
    df_out = clean_table(df_in)

    # "Empty" стовпець порожній -> мав би видалитися
    assert "Empty" not in df_out.columns
    # Маємо 2 стовпці, 2 рядки
    assert df_out.shape == (2, 2)
    # Назви колонок у snake_case
    assert list(df_out.columns) == ["col_1", "col_2"]
    # Порожнє значення -> "N/A"
    assert df_out.iloc[0, 0] == "N/A"
    # Зайві пробіли й перенесення мають зникнути
    assert df_out.iloc[1, 0] == "foo"
    assert df_out.iloc[1, 1] == "Bar"


# ----------------- ІНТЕГРАЦІЙНІ ТЕСТИ (Flask) ----------------- #

@pytest.fixture
def client():
    """
    Створюємо тест-клієнт Flask.
    """
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client

def test_index_route(client):
    resp = client.get('/')
    assert resp.status_code == 200

def test_view_table_csv(client, csv_file_comma):
    data = {
        'file': (csv_file_comma, 'test.csv')
    }
    resp = client.post('/view', data=data, content_type='multipart/form-data')
    assert resp.status_code == 200
    # Перевіряємо наявність HTML-таблиці
    assert b"<table" in resp.data

def test_view_table_unsupported(client):
    data = {
        'file': (io.BytesIO(b'dummy content'), 'somefile.xyz')
    }
    resp = client.post('/view', data=data, content_type='multipart/form-data')
    assert resp.status_code == 400
    assert b"Unsupported file type" in resp.data

def test_save_file_no_session(client):
    resp = client.get('/save')
    assert resp.status_code == 400
    assert b"No data to save!" in resp.data


# ----------------- Фікстура прибирання ----------------- #

@pytest.fixture(autouse=True)
def cleanup_converted_folder():
    """
    Після кожного тесту чистимо 'converted_files/'
    від згенерованих файлів (xlsx тощо).
    """
    yield
    for fname in os.listdir(CONVERTED_FOLDER):
        file_path = os.path.join(CONVERTED_FOLDER, fname)
        if os.path.isfile(file_path):
            os.remove(file_path)
