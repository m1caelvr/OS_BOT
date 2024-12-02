import asyncio
from screeninfo import get_monitors
import pydirectinput as pdi
import pandas as pd
import pyperclip
import pyautogui
import logging
import time
from Class.shared_state import SharedState

monitor = get_monitors()[0]

original_width = 1920
original_height = 1080
new_width = monitor.width
new_height = monitor.height

scale_x = new_width / original_width
scale_y = new_height / original_height

SOURCE_FILE = "PREVENTIVAS.xlsx"

COORDINATES = {
    "CLICK_TO_INSERT_OS": (452, 278),
    "CLICK_OUTSIDE_THE_INPUT": (316, 330),
    "CLICK_TO_ADD_DOC": (1755, 216),
    "CLICK_TO_ADD_FILE": (1677, 155),
    "INSERT_NAME_FILE": (1647, 503),
    "CLICK_INPUT": (1723, 289),
    "CLICK_IN_FILE": (985, 409),
    "CLICK_IN_OK": (1721, 674),
    "CLICK_PROCEDURE_OS": (917, 176),
    "CLICK_INPUT_RADIO_HEADER": (1172, 255),
    "CLICK_ON_LABOR": (609, 181),
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

COORDINATES_2 = {
    "CLICK_TO_INSERT_OS": (273,224),
    "CLICK_OUTSIDE_THE_INPUT": (214,268),
    "CLICK_TO_ADD_DOC": (1220,171),
    "CLICK_TO_ADD_FILE": (1096,128),
    "INSERT_NAME_FILE": (1144,402),
    "CLICK_INPUT": (1200,232),
    "CLICK_IN_FILE": (686,251),
    "CLICK_IN_OK": (1200,515),
    "CLICK_PROCEDURE_OS": (699,145),
    "CLICK_INPUT_RADIO_HEADER": (822,206),
    "CLICK_ON_LABOR": (442,145),
    "CLICK_ADD_LINE": (295,205),
    "INITIAL_SERVICE_DATE": (811,262),
    "FINAL_SERVICE_DATE": (927,257),
    "CLICK_SAVE": (1063,118),
    "SEARCH_OS_STATE": (392,383),
    "CLICK_OS_CONCLUDE": (361,317),
    "CLICK_END_SERVICE": (692,482),
}

logging.basicConfig(
    level=logging.INFO,
    filename="automacao.log",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s",
)

async def log_stop_execution():
    print(f"stop_execution: {SharedState.stop_execution}")

async def sleep(seconds):
    time.sleep(seconds)

async def safe_click(coords):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return
    await asyncio.to_thread(pyautogui.moveTo, *coords)
    await asyncio.to_thread(pyautogui.click)
    await asyncio.sleep(0.5)

async def keyboard_pressed(key):
    pdi.press(key)

async def hotkey(*keys):
    for key in keys[:-1]:
        pdi.keyDown(key)
    pdi.press(keys[-1])
    for key in keys[:-1]:
        pdi.keyUp(key)

async def finalize_line(df, line_index, file_name="PREVENTIVAS.xlsx"):
    try:
        if line_index in df.index:
            df.at[line_index, "Status"] = "Finalizada"
            df.to_excel(file_name, index=False)
            logging.info(f"Planilha atualizada com sucesso: {file_name}")
        else:
            logging.warning(f"Índice {line_index} não encontrado no DataFrame.")
    except Exception as e:
        logging.error(f"Erro ao finalizar linha {line_index}: {e}")

async def insert_os(cell_value):
    await safe_click(COORDINATES_2["CLICK_TO_INSERT_OS"])
    await hotkey("ctrl", "a")
    pyautogui.write(str(cell_value))
    await safe_click(COORDINATES_2["CLICK_OUTSIDE_THE_INPUT"])
    await sleep(4.5)

async def add_doc():
    await safe_click(COORDINATES_2["CLICK_TO_ADD_DOC"])
    await safe_click(COORDINATES_2["CLICK_TO_ADD_FILE"])
    pyperclip.copy(CONSTANTS["FILE_IN_PRISMA_NAME"])
    await safe_click(COORDINATES_2["INSERT_NAME_FILE"])
    await hotkey("ctrl", "v")
    await safe_click(COORDINATES_2["CLICK_INPUT"])
    await sleep(1)
    await safe_click(COORDINATES_2["CLICK_IN_FILE"])
    pyautogui.press("enter")
    await sleep(.5)
    await safe_click(COORDINATES_2["CLICK_IN_OK"])
    await sleep(1)

async def fill_data():
    await safe_click(COORDINATES_2["CLICK_OUTSIDE_THE_INPUT"])
    await sleep(0.5)
    await safe_click(COORDINATES_2["CLICK_PROCEDURE_OS"])
    await sleep(.5)
    await safe_click(COORDINATES_2["CLICK_INPUT_RADIO_HEADER"])
    await sleep(2)
    await safe_click(COORDINATES_2["CLICK_ON_LABOR"])
    await safe_click(COORDINATES_2["CLICK_ADD_LINE"])
    await sleep(1)
    
    pyautogui.write("EQS")
    await sleep(1)
    await keyboard_pressed("tab")
    await keyboard_pressed("tab")
    await sleep(1.5)
    pyperclip.copy(CONSTANTS["INITIAL_DATE"])
    await hotkey("ctrl", "v")
    await sleep(1)
    await keyboard_pressed("tab")
    await sleep(1)
    pyperclip.copy(CONSTANTS["FINAL_DATE"])
    await hotkey("ctrl", "v")
    await sleep(1)
    await keyboard_pressed("tab")
    await sleep(1)
    pyautogui.write(CONSTANTS["HN"])
    await sleep(1)
    await safe_click(COORDINATES_2["CLICK_OUTSIDE_THE_INPUT"])
    await sleep(0.5)

async def end_service():
    await safe_click(COORDINATES_2["CLICK_SAVE"])
    await sleep(4.5)
    await safe_click(COORDINATES_2["CLICK_OUTSIDE_THE_INPUT"])
    await sleep(4.5)
    await safe_click(COORDINATES_2["SEARCH_OS_STATE"])
    await sleep(0.5)
    await safe_click(COORDINATES_2["CLICK_OS_CONCLUDE"])
    await sleep(1.5)
    await safe_click(COORDINATES_2["CLICK_SAVE"])
    await sleep(1)
    await safe_click(COORDINATES_2["CLICK_END_SERVICE"])
    await sleep(4)

async def process_lines(df):
    df["Status"] = df["Status"].astype(str)

    for index, row in df.iterrows():
        if SharedState.stop_execution:
            logging.info("Execução interrompida pelo usuário.")
            break

        cell_value = row.iloc[0]
        status_value = row["Status"]

        logging.info(f"Processando linha {index + 2}: {cell_value} | Status: {status_value}")

        try:
            if status_value != "Finalizada":
                await insert_os(cell_value)
                await add_doc()
                await fill_data()
                await end_service()
                await finalize_line(df, index)
            else:
                logging.info(f"Linha {index + 2} já está finalizada. Pulando...")
        except Exception as e:
            logging.error(f"Erro ao processar linha {index + 2}: {e}")
            print(f"Erro ao processar linha {index + 2}: {e}")

async def start_bot():
    try: 
        logging.info("Iniciando o script de automação...")
        df = pd.read_excel(SOURCE_FILE)

        if "Status" not in df.columns:
            raise ValueError("A coluna 'Status' não foi encontrada no arquivo Excel.")

        logging.info("Arquivo Excel carregado com sucesso.")

        await process_lines(df)
        logging.info("Automação concluída com sucesso.")
    except Exception as e:
        logging.error(f"Erro crítico: {e}")
        print(f"Erro crítico: {e}")
