import os
import time
import asyncio
import logging
import pyautogui
import pyperclip
import pydirectinput as pdi
from src.shared.shared_state import SharedState
from src.bot.constants import CONSTANTS, SOURCE_FILE
from src.bot.constants import COORDINATES_NOTEBOOK_APRENDIZ as COORDINATES


async def sleep(seconds):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return
    time.sleep(seconds)


async def safe_click(coords):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return
    await asyncio.to_thread(pyautogui.moveTo, *coords)
    await asyncio.to_thread(pyautogui.click)
    await asyncio.sleep(0.5)


async def keyboard_pressed(key):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return
    pdi.press(key)


async def hotkey(*keys):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return
    for key in keys[:-1]:
        pdi.keyDown(key)
    pdi.press(keys[-1])
    for key in keys[:-1]:
        pdi.keyUp(key)


async def past_text(text):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return

    pyperclip.copy(text)

    await hotkey("ctrl", "v")


async def insert_os(cell_value):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return

    await safe_click(COORDINATES.CLICK_TO_INSERT_OS)
    await hotkey("ctrl", "a")
    pyautogui.write(str(cell_value))
    await safe_click(COORDINATES.CLICK_OUTSIDE_THE_INPUT)
    await sleep(6)


async def add_doc(site_number=1):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return
    
    async def down_press(site_number):
        for _ in range(site_number):
            if site_number == 1:
                pyautogui.press("down")
                pyautogui.press("up")
            else:
                pyautogui.press("down")

    await safe_click(COORDINATES.CLICK_TO_ADD_DOC)
    await safe_click(COORDINATES.CLICK_TO_ADD_FILE)
    await safe_click(COORDINATES.INSERT_NAME_FILE)
    await past_text(CONSTANTS.FILE_IN_PRISMA_NAME)
    await safe_click(COORDINATES.CLICK_INPUT)
    await sleep(3)
    await safe_click(COORDINATES.CLICK_IN_FILE)
    await down_press(site_number)
    pyautogui.press("enter")
    await sleep(1)
    await safe_click(COORDINATES.CLICK_IN_OK)
    await safe_click(COORDINATES.CLICK_IN_OK)
    await sleep(2)



async def fill_data():
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return

    await safe_click(COORDINATES.CLICK_OUTSIDE_THE_INPUT)
    await sleep(0.5)
    await safe_click(COORDINATES.CLICK_PROCEDURE_OS)
    await sleep(0.5)
    await safe_click(COORDINATES.CLICK_INPUT_RADIO_HEADER)
    await sleep(5)
    await safe_click(COORDINATES.CLICK_ON_LABOR)
    await safe_click(COORDINATES.CLICK_ADD_LINE)
    await sleep(1)

    pyautogui.write("EQS")
    await sleep(1)
    await keyboard_pressed("tab")
    await keyboard_pressed("tab")
    await sleep(1.5)
    await past_text(CONSTANTS.INITIAL_DATE)
    await sleep(1)
    await keyboard_pressed("tab")
    await sleep(1)
    await past_text(CONSTANTS.FINAL_DATE)
    await sleep(1)
    await keyboard_pressed("tab")
    await sleep(1)
    pyautogui.write(CONSTANTS.HN)
    await sleep(1)
    await safe_click(COORDINATES.CLICK_OUTSIDE_THE_INPUT)
    await sleep(0.5)


async def end_service():
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return

    await safe_click(COORDINATES.CLICK_SAVE)
    await sleep(6.5)
    await safe_click(COORDINATES.CLICK_OUTSIDE_THE_INPUT)
    await sleep(7)
    await safe_click(COORDINATES.SEARCH_OS_STATE)
    await sleep(1)
    await safe_click(COORDINATES.CLICK_OS_CONCLUDE)
    await sleep(1.5)
    await safe_click(COORDINATES.CLICK_SAVE)
    await sleep(1)
    await safe_click(COORDINATES.CLICK_END_SERVICE)
    await safe_click(COORDINATES.CLICK_END_SERVICE)
    await safe_click(COORDINATES.CLICK_END_SERVICE)
    await sleep(5.5)


async def finalize_line(df, line_index, increment, file_name=SOURCE_FILE):
    if SharedState.stop_execution:
        logging.info("Execução interrompida. Cancelando safe_click.")
        return

    try:
        if line_index in df.index:
            df.at[line_index, "Status"] = "Finalizada"
            df.to_excel(file_name, index=False)
            await increment()
            logging.info(f"Planilha atualizada com sucesso: {file_name}")
        else:
            logging.warning(f"Índice {line_index} não encontrado no DataFrame.")
    except Exception as e:
        logging.error(f"Erro ao finalizar linha {line_index}: {e}")
