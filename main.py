# main.py

import sys
import os
import ctypes

from application.execute_refresh_uc import execute_refresh
from infrastructure.config_loader import ConfigLoader
from infrastructure.logger_service import LoggerService
from infrastructure.scheduler_service import SchedulerService

def main():
    # --- CHECK DE SEGURIDAD (FASE 1) ---
    if not os.path.exists(".env"):
        ctypes.windll.user32.MessageBoxW(
            0, 
            "⚠️ Error Crítico: No se encontró el archivo de licencia o configuración (.env).\n\n"
            "Por favor, contacta con Shohan para obtener tu copia personalizada de Pivoty.", 
            "Pivoty - Error de Activación", 
            0x10 | 0x0  # Icono de Error + Botón OK
        )
        sys.exit(1)

    config = ConfigLoader()

    logger = LoggerService(
        log_level=config.get("LOG_LEVEL", "INFO"),
        log_dir=config.get("LOG_DIR", "logs"),
        log_name=config.get("LOG_FILE", "pivoty.log")
    ).get_logger()

    if "--scheduler" in sys.argv:
        scheduler = SchedulerService(
            config=config,
            logger=logger,
            execute_fn=execute_refresh
        )

        scheduler.start()
    elif "--refresh" in sys.argv:
        execute_refresh()
    else:
        from gui.app import run
        run()

if __name__ == "__main__":
    main()
