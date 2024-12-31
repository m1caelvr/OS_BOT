import asyncio
import logging
from src.bot.automation import (
    insert_os,
    add_doc,
    fill_data,
    end_service,
    finalize_line,
)
from src.shared.shared_state import SharedState
from src.bot.constants import CONSTANTS

async def process_lines(df, increment):
    df["Status"] = df["Status"].astype(str)
    df["Denominacao_Site"] = df["Denominacao_Site"].astype(str)

    site_mapping = {}
    current_site_index = 1

    for index, row in df.iterrows():
        if SharedState.stop_execution:
            logging.info("Execução interrompida pelo usuário.")
            break

        cell_value = row["N_OS"]
        status_value = row["Status"]
        site_name = row["Denominacao_Site"]

        logging.info(
            f"Processando linha {index + 2}: N_OS={cell_value} | "
            f"Status={status_value} | Site={site_name}"
        )

        if site_name not in site_mapping:
            site_mapping[site_name] = current_site_index
            current_site_index += 1
        site_index = site_mapping[site_name]

        try:
            if status_value != "Finalizada":
                await insert_os(cell_value)
                await add_doc(site_index)
                await fill_data()
                await end_service()
                await finalize_line(df, index, increment)
            else:
                logging.info(f"Linha {index + 2} já está finalizada. Pulando...")
        except Exception as e:
            logging.error(f"Erro ao processar linha {index + 2}: {e}")


async def start_bot(increment, df):
    try:
        logging.info("Iniciando o script de automação...")
        await process_lines(df, increment)
        logging.info("Automação concluída com sucesso.")
    except Exception as e:
        logging.error(f"Erro crítico: {e}")
