import asyncio
import flet as ft
import logging
import tracemalloc
from Class.shared_state import SharedState
from Bot import start_bot

tracemalloc.start()

class ScriptController:
    def __init__(self, page):
        self.page = page
        self.script_task = None
        self.start_button = None
        self.stop_button = None
        self.is_running = False

    async def run_script(self):
        try:
            logging.info("Iniciando o script de automação...")
            await start_bot()
        except asyncio.CancelledError:
            logging.info("Script foi interrompido.")
            print("Script foi interrompido.")
        except Exception as e:
            logging.error(f"Erro crítico: {e}")
            print(f"Erro crítico: {e}")
            
            snapshot = tracemalloc.take_snapshot()
            top_stats = snapshot.statistics('lineno')
            print("[ Top 10 Memory Allocation Lines ]")
            for stat in top_stats[:10]:
                print(stat)
        finally:
            await self.reset_ui()

    async def toggle_script(self, e, page):
        if not self.is_running:
            SharedState.stop_execution = False
            self.is_running = True
            self.start_button.disabled = True
            self.stop_button.disabled = False
            page.update()

            await asyncio.sleep(.1)

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

            await asyncio.sleep(.1)

            if self.script_task and not self.script_task.done():
                self.script_task.cancel()
                try:
                    await self.script_task
                except asyncio.CancelledError:
                    print("Tarefa cancelada com sucesso")

    async def reset_ui(self):
        self.is_running = False
        self.start_button.disabled = False
        self.stop_button.disabled = True
        self.page.update()

def main(page: ft.Page):
    page.window.width = 250
    page.window.height = 150
    page.title = "OS_BOT"
    page.horizontal_alignment = "center"
    page.window.always_on_top = True

    controller = ScriptController(page)

    async def toggle_script_wrapper(e):
        await asyncio.sleep(.1)
        await controller.toggle_script(e, page)

    controller.start_button = ft.ElevatedButton(
        "Iniciar BOT",
        on_click=toggle_script_wrapper
    )
    controller.stop_button = ft.ElevatedButton(
        "Parar BOT",
        on_click=toggle_script_wrapper
    )

    controller.stop_button.disabled = True

    page.add(controller.start_button, controller.stop_button)
    page.update()

    page.on_close = lambda _: tracemalloc.stop()

ft.app(target=main)
