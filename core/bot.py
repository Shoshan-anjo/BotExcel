# bot.py
import os
import json
import logging
from datetime import datetime
from pathlib import Path
from dotenv import dotenv_values
from excel_manager import ExcelManager
from mail_manager import MailManager
from scheduler import Scheduler

# -------------------------
# Configuración
# -------------------------
CONFIG_PATH = Path("config/excels.json")
LOGS_PATH = Path("logs")
LOGS_PATH.mkdir(exist_ok=True)

logging.basicConfig(
    filename=LOGS_PATH / f"{datetime.now().strftime('%Y-%m-%d')}.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

env = dotenv_values(".env")

# -------------------------
# Inicializar módulos
# -------------------------
excel_manager = ExcelManager(CONFIG_PATH)
mail_manager = MailManager(env)
scheduler = Scheduler()

# -------------------------
# Función de ejecución principal
# -------------------------
def run_task():
    """
    Ejecuta la actualización de todos los Excel activos.
    """
    excels = excel_manager.load_excels()
    for excel in excels:
        if not excel["activo"]:
            continue

        path = excel["path"]
        backup_path = excel.get("backup")
        try:
            # 1️⃣ Validación avanzada
            if not excel_manager.validate_excel(path):
                logging.error(f"Archivo inválido: {path}")
                continue

            # 2️⃣ Lógica de actualización (placeholder)
            excel_manager.update_excel(path, backup_path)

            # 3️⃣ Validación de filas
            if excel.get("validate_rows", False):
                if not excel_manager.validate_rows(path, excel.get("min_rows", 1)):
                    logging.warning(f"Validación de filas fallida: {path}")
                    continue

            logging.info(f"Excel actualizado correctamente: {path}")

            # 4️⃣ Enviar correo si está habilitado
            if mail_manager.enabled():
                mail_manager.send_email(excel)

        except Exception as e:
            logging.exception(f"Error actualizando {path}: {e}")

# -------------------------
# Programar ejecución según horarios
# -------------------------
def schedule_tasks():
    excels = excel_manager.load_excels()
    for excel in excels:
        for horario in excel.get("horarios", []):
            scheduler.schedule_task(horario, run_task)

# -------------------------
# Modo CLI
# -------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="BotExcel Headless Runner")
    parser.add_argument("--run", action="store_true", help="Ejecutar actualizaciones sin GUI")
    args = parser.parse_args()

    if args.run:
        schedule_tasks()
        scheduler.run_forever()
    else:
        from gui.app import start_gui
        start_gui()
