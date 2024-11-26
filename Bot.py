import pydirectinput as pdi
import pandas as pd
import pyperclip
import pyautogui
import keyboard
import logging
import time
import os

COORDINATES = {
    "CLICK_TO_INSERT_OS_N": (452, 278),
    "CLICK_OUTSIDE_THE_INPUT_N": (316, 330),
    "CLICK_TO_ADD_DOC_N": (1755, 216),
    "CLICK_TO_ADD_FILE_N": (1677, 155),
    "INSERT_NAME_FILE_N": (1647, 503),
    "CLICK_INPUT_N": (1723, 289),
    "CLICK_IN_FILE_N": (985, 409),
    "CLICK_IN_OK_N": (1721, 674),
    "CLICK_PROCEDURE_OS_N": (917, 176),
    "CLICK_INPUT_RADIO_HEADER_N": (1172, 255),
    "CLICK_ON_LABOR_N": (609, 181),
    "CLICK_ADD_LINE": (435, 253),
    "INITIAL_SERVICE_DATE": (1055, 318),
    "FINAL_SERVICE_DATE": (1212, 319),
    "CLICK_SAVE": (1533, 152),
    "SEARCH_OS_STATE": (573, 481),
    "CLICK_OS_CONCLUDE": (495, 398),
    "CLICK_END_SERVICE": (1004, 665),
}

CONSTANTS = {
    "FILE_IN_PRISMA_NAME": "RELATÓRIO",
    "INITIAL_DATE": "01/10/2024 08:00",
    "FINAL_DATE": "01/10/2024 09:00",
    "HN": "HN",
}

logging.basicConfig(
    level=logging.INFO,
    filename="automacao.log",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s",
)

SOURCE_FILE = "PREVENTIVAS.xlsx"


def sleep(seconds):
    time.sleep(seconds)


def safe_click(coords):
    pdi.moveTo(*coords)
    pdi.click()
    sleep(0.5)


def keyboard_pressed(key):
    pdi.press(key)


def hotkey(*keys):
    for key in keys[:-1]:
        pdi.keyDown(key)
    pdi.press(keys[-1])
    for key in keys[:-1]:
        pdi.keyUp(key)


def wait_until_found(image, timeout=10):
    start_time = time.time()
    while time.time() - start_time < timeout:
        location = pyautogui.locateOnScreen(image, confidence=0.9)
        if location:
            return location
        sleep(0.5)
    raise TimeoutError(f"Imagem {image} não encontrada em {timeout} segundos.")


def insert_os(cell_value):
    safe_click(COORDINATES["CLICK_TO_INSERT_OS_N"])
    hotkey("ctrl", "a")
    pyautogui.write(str(cell_value))
    safe_click(COORDINATES["CLICK_OUTSIDE_THE_INPUT_N"])
    sleep(4.5)


def add_doc():
    safe_click(COORDINATES["CLICK_TO_ADD_DOC_N"])
    safe_click(COORDINATES["CLICK_TO_ADD_FILE_N"])
    pyperclip.copy(CONSTANTS["FILE_IN_PRISMA_NAME"])
    safe_click(COORDINATES["INSERT_NAME_FILE_N"])
    hotkey("ctrl", "v")
    safe_click(COORDINATES["CLICK_INPUT_N"])
    safe_click(COORDINATES["CLICK_IN_FILE_N"])
    safe_click(COORDINATES["CLICK_IN_OK_N"])
    sleep(1)


def fill_data():
    safe_click(COORDINATES["CLICK_OUTSIDE_THE_INPUT_N"])
    sleep(0.5)
    safe_click(COORDINATES["CLICK_PROCEDURE_OS_N"])
    sleep(1)
    safe_click(COORDINATES["CLICK_INPUT_RADIO_HEADER_N"])
    sleep(2)
    safe_click(COORDINATES["CLICK_ON_LABOR_N"])
    safe_click(COORDINATES["CLICK_ADD_LINE"])

    sleep(1)
    pyautogui.write("EQS")
    sleep(1)
    keyboard_pressed("tab")
    keyboard_pressed("tab")
    sleep(1.5)
    pyperclip.copy(CONSTANTS["INITIAL_DATE"])
    hotkey("ctrl", "v")
    sleep(1)
    keyboard_pressed("tab")
    sleep(1)
    pyperclip.copy(CONSTANTS["FINAL_DATE"])
    hotkey("ctrl", "v")
    sleep(1)
    keyboard_pressed("tab")
    sleep(1.5)
    pyautogui.write(CONSTANTS["HN"])
    sleep(1)
    safe_click(COORDINATES["CLICK_OUTSIDE_THE_INPUT_N"])
    sleep(0.5)


def process_lines(df):
    df["Status"] = df["Status"].astype(str)

    for index, row in df.iterrows():
        if keyboard.is_pressed("esc"):
            logging.info("Script interrompido pelo usuário.")
            print("Script interrompido pelo usuário!")
            break

        cell_value = row.iloc[0]
        status_value = row["Status"]

        logging.info(
            f"Processando linha {index + 1}: {cell_value} | Status: {status_value}"
        )

        try:
            if status_value != "Finalizada":
                insert_os(cell_value)
                add_doc()
                fill_data()

                safe_click(COORDINATES["CLICK_SAVE"])
                sleep(4)
                safe_click(COORDINATES["CLICK_OUTSIDE_THE_INPUT_N"])
                sleep(4.5)
                safe_click(COORDINATES["SEARCH_OS_STATE"])
                sleep(0.5)
                safe_click(COORDINATES["CLICK_OS_CONCLUDE"])
                sleep(1.5)
                safe_click(COORDINATES["CLICK_SAVE"])
                sleep(1)
                safe_click(COORDINATES["CLICK_END_SERVICE"])
                sleep(4)

                df.at[index, "Status"] = "Finalizada"
                logging.info(f"Linha {index + 1} - Status atualizado para: Finalizada")
                print(f"Linha {index + 1} - Status atualizado para: Finalizada")
            else:
                logging.info(f"Linha {index + 1} já está finalizada. Pulando...")

        except Exception as e:
            logging.error(f"Erro ao processar linha {index + 1}: {e}")
            print(f"Erro ao processar linha {index + 1}: {e}")

    df.to_excel(SOURCE_FILE, index=False)
    logging.info("Arquivo Excel atualizado com o status das OSs.")


if __name__ == "__main__":
    try:
        logging.info("Iniciando o script de automação...")
        print("Pressione 'ESC' para interromper o script.")

        df = pd.read_excel(SOURCE_FILE)

        if "Status" not in df.columns:
            raise ValueError("A coluna 'Status' não foi encontrada no arquivo Excel.")

        logging.info("Arquivo Excel carregado com sucesso.")

        process_lines(df)
        logging.info("Automação concluída com sucesso.")
    except Exception as e:
        logging.error(f"Erro crítico: {e}")
        print(f"Erro crítico: {e}")
