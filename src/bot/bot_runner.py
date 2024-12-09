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


async def process_lines(df, increment):
    df["Status"] = df["Status"].astype(str)

    for index, row in df.iterrows():
        if SharedState.stop_execution:
            logging.info("Execução interrompida pelo usuário.")
            break

        cell_value = row.iloc[0]
        status_value = row["Status"]

        logging.info(
            f"Processando linha {index + 2}: {cell_value} | Status: {status_value}"
        )

        try:
            if status_value != "Finalizada":
                await insert_os(cell_value)
                await add_doc()
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
