import asyncio
import flet as ft
import logging
import tracemalloc

from screeninfo import get_monitors
from Class.shared_state import SharedState

from src.Bot import start_bot

tracemalloc.start()


class ScriptController:
    def __init__(self, page):
        self.page = page
        self.script_task = None
        self.start_button = None
        self.stop_button = None
        self.is_running = False
        self.os_count_label = None

    async def run_script(self):
        try:
            logging.info("Iniciando o script de automação...")
            await start_bot(self.increment_made_consecutively)
        except asyncio.CancelledError:
            logging.info("Script foi interrompido.")
            print("Script foi interrompido.")
        except Exception as e:
            logging.error(f"Erro crítico: {e}")
            print(f"Erro crítico: {e}")

            snapshot = tracemalloc.take_snapshot()
            top_stats = snapshot.statistics("lineno")
            print("[ Top 10 Memory Allocation Lines ]")
            for stat in top_stats[:10]:
                print(stat)
        finally:
            await self.reset_ui()

    async def toggle_script(self, e, page):
        if not self.is_running:
            SharedState.stop_execution = False
            SharedState.made_consecutively = 0
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
                    print("Tarefa cancelada com sucesso")

    async def increment_made_consecutively(self):
        SharedState.made_consecutively += 1
        self.update_os_count()

    def update_os_count(self):
        self.os_count_label.value = (
            f"Quantidade feita consecutiva: {SharedState.made_consecutively}"
        )
        self.page.update()
        print(f"Quantidade feita consecutiva: {SharedState.made_consecutively}")

    async def reset_ui(self):
        self.is_running = False
        self.start_button.disabled = False
        self.stop_button.disabled = True
        SharedState.made_consecutively = 0
        self.update_os_count()
        self.page.update()


def main(page: ft.Page):
    window_width = 300
    window_height = 200
    page.window.width = window_width
    page.window.height = window_height

    monitor = get_monitors()[0]
    screen_width = monitor.width
    screen_height = monitor.height

    page.window.left = screen_width / 3
    page.window.top = screen_height / 2

    page.title = "OS_BOT"
    page.horizontal_alignment = "center"
    page.window.always_on_top = True

    controller = ScriptController(page)

    async def toggle_script_wrapper(e):
        await asyncio.sleep(0.1)
        await controller.toggle_script(e, page)

    controller.start_button = ft.ElevatedButton(
        "Iniciar BOT", on_click=toggle_script_wrapper
    )
    controller.stop_button = ft.ElevatedButton(
        "Parar BOT",
        on_click=toggle_script_wrapper,
    )

    controller.stop_button.disabled = True

    controller.os_count_label = ft.Text("Quantidade feita consecutiva: 0")

    page.add(controller.start_button, controller.stop_button, controller.os_count_label)
    page.update()

    page.on_close = lambda _: tracemalloc.stop()


ft.app(target=main)
