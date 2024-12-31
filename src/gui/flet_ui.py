import os
import asyncio
import logging
import tracemalloc
import pdfplumber
import re

import flet as ft
import pandas as pd
from screeninfo import get_monitors
from datetime import datetime

from src.bot.bot_runner import start_bot
from src.bot.constants import CONSTANTS, SOURCE_FILE
from src.shared.shared_state import SharedState
from src.shared.config_manager import (
    load_config,
    save_config,
)

log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_file = os.path.join(log_dir, "automacao.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
)

tracemalloc.start()


class ScriptController:
    def __init__(self, page):
        self.page = page
        self.script_task = None
        self.start_button = None
        self.stop_button = None
        self.is_running = False
        self.os_count_label = None
        self.os_count_restant = None
        self.df = pd.read_excel(SOURCE_FILE)
        self.config = load_config()

    async def run_script(self):
        try:
            logging.info("Iniciando o script de automação...")
            await start_bot(self.increment_made_consecutively, self.df)
        except asyncio.CancelledError:
            logging.info("Script foi interrompido.")
        except Exception as e:
            logging.error(f"Erro crítico: {e}")

            snapshot = tracemalloc.take_snapshot()
            top_stats = snapshot.statistics("lineno")
            logging.info("[ Top 10 Memory Allocation Lines ]")
            for stat in top_stats[:10]:
                logging.info(stat)
        finally:
            await self.reset_ui()

    async def toggle_script(self, e, page):
        if not self.is_running:
            SharedState.stop_execution = False
            SharedState.made_consecutively = 0
            SharedState.os_restants = self.lines_for_finalize()
            self.update_os_count()
            self.is_running = True
            self.start_button.disabled = True
            self.stop_button.disabled = False
            page.update()

            await asyncio.sleep(0.1)

            if self.script_task and not self.script_task.done():
                self.script_task.cancel()
                try:
                    await self.script_task
                except asyncio.CancelledError:
                    pass

            self.script_task = asyncio.create_task(self.run_script())
        else:
            SharedState.stop_execution = True
            self.is_running = False
            self.start_button.disabled = False
            self.stop_button.disabled = True
            page.update()

            await asyncio.sleep(0.1)

            if self.script_task and not self.script_task.done():
                self.script_task.cancel()
                try:
                    await self.script_task
                except asyncio.CancelledError:
                    logging.info("Tarefa cancelada com sucesso")

    async def increment_made_consecutively(self):
        SharedState.made_consecutively += 1
        SharedState.os_restants -= 1
        self.update_os_count()

    def update_os_count(self):
        self.os_count_label.value = (
            f"Quantidade feita consecutiva: {SharedState.made_consecutively}"
        )
        self.os_count_restant.value = (
            f"Preventivas restantes {SharedState.os_restants} de {len(self.df)}"
        )
        self.page.update()

    async def reset_ui(self):
        self.is_running = False
        self.start_button.disabled = False
        self.stop_button.disabled = True
        SharedState.made_consecutively = 0
        SharedState.os_restants = self.lines_for_finalize()
        self.update_os_count()
        self.page.update()

    def lines_for_finalize(self):
        df = self.df

        if "Status" not in df.columns:
            df["Status"] = ""
            df.to_excel(SOURCE_FILE, index=False)
            logging.info("Coluna 'Status' criada.")

        for_finalize = df[df["Status"] != "Finalizada"]

        return len(for_finalize)


def validate_datetime(value):
    try:
        datetime.strptime(value, "%d/%m/%Y %H:%M")
        return True
    except ValueError:
        return False


def main(page: ft.Page):
    window_width = 400
    window_height = 330
    page.window.width = window_width
    page.window.height = window_height

    monitor = get_monitors()[0]
    screen_width = monitor.width
    screen_height = monitor.height

    page.window.left = window_width / 3
    page.window.top = screen_height - window_height * 1.7

    page.title = "OS_BOT"
    page.horizontal_alignment = "center"
    page.window.always_on_top = True

    # page.theme_mode = ft.ThemeMode.LIGHT

    controller = ScriptController(page)
    config = controller.config

    async def toggle_script_wrapper(e):
        await asyncio.sleep(0.1)
        await controller.toggle_script(e, page)

    initial_date_input = ft.TextField(
        label="Data Inicial",
        value=config.get("INITIAL_DATE", ""),
        hint_text="dd/mm/aaaa hh:mm",
        width=300,
    )
    final_date_input = ft.TextField(
        label="Data Final",
        value=config.get("FINAL_DATE", ""),
        hint_text="dd/mm/aaaa hh:mm",
        width=300,
    )

    insert_date = ft.Row(
        [
            ft.Container(
                content=initial_date_input,
                expand=True,
            ),
            ft.Container(
                content=final_date_input,
                expand=True,
            ),
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=10,
    )

    def save_config_handler(e):
        initial_date = initial_date_input.value
        final_date = final_date_input.value

        if not validate_datetime(initial_date):
            page.snack_bar = ft.SnackBar(ft.Text("Data Inicial inválida!"))
            page.snack_bar.open = True
            page.update()
            return

        if not validate_datetime(final_date):
            page.snack_bar = ft.SnackBar(ft.Text("Data Final inválida!"))
            page.snack_bar.open = True
            page.update()
            return

        config["INITIAL_DATE"] = initial_date
        config["FINAL_DATE"] = final_date
        save_config(config, CONSTANTS)

        page.snack_bar = ft.SnackBar(ft.Text("Configurações salvas com sucesso!"))
        page.snack_bar.open = True
        page.update()

    save_button = ft.ElevatedButton(
        "Salvar Configurações", on_click=save_config_handler
    )

    controller.start_button = ft.ElevatedButton(
        "Iniciar BOT", on_click=toggle_script_wrapper
    )
    controller.stop_button = ft.ElevatedButton(
        "Parar BOT",
        on_click=toggle_script_wrapper,
    )
    controller.stop_button.disabled = True

    button_controller = ft.Row(
        [
            ft.Container(
                content=controller.start_button,
                expand=True,
            ),
            ft.Container(
                content=controller.stop_button,
                expand=True,
            ),
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=10,
    )

    for_finalize = controller.lines_for_finalize()
    total_lines = len(controller.df)

    controller.os_count_label = ft.Text("Quantidade feita consecutiva: 0")
    controller.os_count_restant = ft.Text(
        f"Preventivas restantes {for_finalize} de {total_lines}"
    )

    def handle_upload_planilha(e: ft.FilePickerResultEvent):
        if e.files:
            file = e.files[0]
            directory = os.path.join("data", "PLANILHA_OS")
            new_path = os.path.join(directory, "PREVENTIVAS.xlsx")

            try:
                if os.path.exists(directory):
                    for existing_file in os.listdir(directory):
                        existing_file_path = os.path.join(directory, existing_file)
                        if os.path.isfile(existing_file_path):
                            os.remove(existing_file_path)
                else:
                    os.makedirs(directory)

                with open(file.path, "rb") as src, open(new_path, "wb") as dest:
                    dest.write(src.read())

                df = pd.read_excel(new_path)

                if "N_OS" not in df.columns:
                    page.snack_bar = ft.SnackBar(
                        ft.Text("Erro: Coluna 'N_OS' não encontrada na planilha.")
                    )
                    page.snack_bar.open = True
                    page.update()
                    return

                df = df[["N_OS"]]
                df.to_excel(new_path, index=False)

                print(f"Planilha: {len(controller.df)}")

                controller.df = df
                print(f"Planilha atualizada com sucesso. {len(controller.df)}")

                if "Status" not in df.columns:
                    df["Status"] = ""
                    df.to_excel(SOURCE_FILE, index=False)
                    logging.info("Coluna 'Status' criada.")

                for_finalize = len(df[df["Status"] != "Finalizada"])
                total_lines = len(df)

                controller.os_count_restant.value = (
                    f"Preventivas restantes {for_finalize} de {total_lines}"
                )
                controller.os_count_restant.update()

                page.snack_bar = ft.SnackBar(
                    ft.Text("Arquivo de planilha atualizado com sucesso!")
                )
                page.snack_bar.open = True
                page.update()

            except Exception as ex:
                page.snack_bar = ft.SnackBar(
                    ft.Text(f"Erro ao processar a planilha: {ex}")
                )
                page.snack_bar.open = True
                page.update()

    def handle_upload_relatorio(e: ft.FilePickerResultEvent):
        if e.files:
            file = e.files[0]
            dir_path = os.path.join("data", "RECORD")
            new_path = os.path.join(dir_path, "Relatório.pdf")

            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

            try:
                for existing_file in os.listdir(dir_path):
                    if existing_file.endswith(".pdf"):
                        os.remove(os.path.join(dir_path, existing_file))

                with open(file.path, "rb") as src, open(new_path, "wb") as dest:
                    dest.write(src.read())

                def extract_date_from_pdf(pdf_path, page_number=0):
                    try:
                        with pdfplumber.open(pdf_path) as pdf:
                            page = pdf.pages[page_number]
                            text = page.extract_text()

                            if not text:
                                return "Texto não encontrado na página especificada."

                            match = re.search(
                                r"FIM:\s*(?:\n|.*?)(\d{2}/\d{2}/\d{4} \d{2}:\d{2})", text, re.DOTALL
                            )

                            if match:
                                return match.group(1)
                            else:
                                return "Formato esperado de 'FIM' não encontrado."

                    except Exception as e:
                        return f"Erro ao processar o PDF: {e}"

                data_hora = extract_date_from_pdf(new_path)

                if isinstance(data_hora, str) and not data_hora.startswith("Erro"):
                    data = data_hora.split(" ")[0]

                    initial_date_value = f"{data} 08:00"
                    final_date_value = f"{data} 09:00"

                    initial_date_input.value = initial_date_value
                    final_date_input.value = final_date_value

                    page.snack_bar = ft.SnackBar(
                        ft.Text("Arquivo de relatório e data atualizados com sucesso!")
                    )
                else:
                    page.snack_bar = ft.SnackBar(
                        ft.Text(f"Erro ao processar o relatório: {data_hora}")
                    )

            except Exception as ex:
                page.snack_bar = ft.SnackBar(
                    ft.Text(f"Erro ao salvar o relatório: {ex}")
                )
            page.snack_bar.open = True
            page.update()


    upload_planilha_button = ft.FilePicker(on_result=handle_upload_planilha)
    upload_relatorio_button = ft.FilePicker(on_result=handle_upload_relatorio)

    page.overlay.extend([upload_planilha_button, upload_relatorio_button])

    upload_buttons = ft.Row(
        [
            ft.Container(
                content=ft.ElevatedButton(
                    "Nova Planilha",
                    on_click=lambda _: upload_planilha_button.pick_files(),
                ),
                expand=True,
            ),
            ft.Container(
                content=ft.ElevatedButton(
                    "Novo Relatório",
                    on_click=lambda _: upload_relatorio_button.pick_files(),
                ),
                expand=True,
            ),
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=10,
    )

    def reload_data(e):
        try:
            controller.df = pd.read_excel(SOURCE_FILE)

            if "Status" not in controller.df.columns:
                controller.df["Status"] = ""
                controller.df.to_excel(SOURCE_FILE, index=False)
                logging.info("Coluna 'Status' criada.")

            for_finalize = controller.lines_for_finalize()
            total_lines = len(controller.df)

            controller.os_count_label.value = (
                f"Quantidade feita consecutiva: {SharedState.made_consecutively}"
            )
            controller.os_count_restant.value = (
                f"Preventivas restantes {for_finalize} de {total_lines}"
            )

            page.snack_bar = ft.SnackBar(ft.Text("Dados recarregados com sucesso!"))
            page.snack_bar.open = True
            page.update()
        except Exception as ex:
            logging.error(f"Erro ao recarregar os dados: {ex}")
            page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao recarregar os dados: {ex}"))
            page.snack_bar.open = True
            page.update()

    reload_button = ft.ElevatedButton("Recarregar Dados", on_click=reload_data)

    page.add(
        insert_date,
        save_button,
        upload_buttons,
        button_controller,
        controller.os_count_label,
        controller.os_count_restant,
        reload_button,
    )

    page.update()
    page.on_close = lambda _: tracemalloc.stop()
