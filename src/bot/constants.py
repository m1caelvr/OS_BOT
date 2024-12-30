import logging
import os
from src.shared.config_manager import load_config

PLANILHA_DIR = "./data/PLANILHA_OS"


def find_excel_file(directory):
    for file in os.listdir(directory):
        if file.endswith((".xlsx", ".xls")):
            return os.path.join(directory, file)
    raise FileNotFoundError(
        "Nenhuma planilha foi encontrada no diretório especificado."
    )


SOURCE_FILE = find_excel_file(PLANILHA_DIR)
logging.info(f"Planilha encontrada: {SOURCE_FILE}")

# 26/11/2024 08:00
# 26/11/2024 09:00


class CONSTANTS:
    FILE_IN_PRISMA_NAME = "RELATÓRIO"
    HN = "HN"

    config = load_config()
    INITIAL_DATE = config.get("INITIAL_DATE")
    FINAL_DATE = config.get("FINAL_DATE")


class COORDINATES_NOTEBOOK_APRENDIZ:
    CLICK_TO_INSERT_OS = [452, 278]
    CLICK_OUTSIDE_THE_INPUT = [316, 330]
    CLICK_TO_ADD_DOC = [1755, 216]
    CLICK_TO_ADD_FILE = [1677, 155]
    INSERT_NAME_FILE = [1647, 503]
    CLICK_INPUT = [1723, 289]
    CLICK_IN_FILE = [480, 178]
    CLICK_IN_OK = [1718, 644]
    CLICK_PROCEDURE_OS = [917, 176]
    CLICK_INPUT_RADIO_HEADER = [1172, 255]
    CLICK_ON_LABOR = [609, 181]
    CLICK_ADD_LINE = [435, 253]
    INITIAL_SERVICE_DATE = [1055, 318]
    FINAL_SERVICE_DATE = [1212, 319]
    CLICK_SAVE = [1533, 152]
    SEARCH_OS_STATE = [573, 481]
    CLICK_OS_CONCLUDE = [495, 398]
    CLICK_END_SERVICE = [1067, 671]


# class COORDINATES_NOTEBOOK_APRENDIZ:
#     CLICK_TO_INSERT_OS = [452, 278]
#     CLICK_OUTSIDE_THE_INPUT = [316, 330]
#     CLICK_TO_ADD_DOC = [1755, 216]
#     CLICK_TO_ADD_FILE = [1677, 155]
#     INSERT_NAME_FILE = [1647, 503]
#     CLICK_INPUT = [1723, 289]
#     CLICK_IN_FILE = [480, 178]
#     CLICK_IN_OK = [1718, 644]
#     CLICK_PROCEDURE_OS = [917, 176]
#     CLICK_INPUT_RADIO_HEADER = [1172, 255]
#     CLICK_ON_LABOR = [609, 181]
#     CLICK_ADD_LINE = [435, 253]
#     INITIAL_SERVICE_DATE = [1055, 318]
#     FINAL_SERVICE_DATE = [1212, 319]
#     CLICK_SAVE = [1533, 152]
#     SEARCH_OS_STATE = [573, 481]
#     CLICK_OS_CONCLUDE = [495, 398]
#     CLICK_END_SERVICE = [1004, 665]


class COORDINATES_NOTEBOOK_ANTONIO:
    CLICK_TO_INSERT_OS = [273, 224]
    CLICK_OUTSIDE_THE_INPUT = [214, 268]
    CLICK_TO_ADD_DOC = [1220, 171]
    CLICK_TO_ADD_FILE = [1096, 128]
    INSERT_NAME_FILE = [1144, 402]
    CLICK_INPUT = [1200, 232]
    CLICK_IN_FILE = [686, 251]
    CLICK_IN_OK = [1200, 515]
    CLICK_PROCEDURE_OS = [699, 145]
    CLICK_INPUT_RADIO_HEADER = [822, 206]
    CLICK_ON_LABOR = [442, 145]
    CLICK_ADD_LINE = [295, 205]
    INITIAL_SERVICE_DATE = [811, 262]
    FINAL_SERVICE_DATE = [927, 257]
    CLICK_SAVE = [1063, 118]
    SEARCH_OS_STATE = [392, 383]
    CLICK_OS_CONCLUDE = [361, 317]
    CLICK_END_SERVICE = [692, 482]
