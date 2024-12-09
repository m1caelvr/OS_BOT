import pandas as pd


def load_excel_file(file_path):
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        raise Exception(f"Arquivo não encontrado: {file_path}")
