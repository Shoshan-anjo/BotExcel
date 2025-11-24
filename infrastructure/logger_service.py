# infrastructure/logger_service.py

import logging
import os
from logging.handlers import TimedRotatingFileHandler


class LoggerService:
    """
    Configura y entrega un logger con rotación diaria de archivos.
    """

    def __init__(self, config_loader):
        self.config = config_loader

    def get_logger(self):
        log_dir = self.config.get("LOG_DIR", "logs")
        log_file = self.config.get("LOG_FILE", "botexcel.log")
        level_name = (self.config.get("LOG_LEVEL", "INFO") or "INFO").upper()

        level = getattr(logging, level_name, logging.INFO)

        os.makedirs(log_dir, exist_ok=True)

        logger = logging.getLogger("BotExcel")

        # Evitar agregar handlers duplicados en múltiples ejecuciones
        if logger.handlers:
            logger.setLevel(level)
            return logger

        logger.setLevel(level)

        log_path = os.path.join(log_dir, log_file)

        file_handler = TimedRotatingFileHandler(
            log_path,
            when="midnight",
            backupCount=7,
            encoding="utf-8",
        )

        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s"
        )
        file_handler.setFormatter(formatter)

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        return logger
