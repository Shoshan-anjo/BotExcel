import os
import logging
from logging.handlers import TimedRotatingFileHandler


class LoggerService:
    """
    Inicializa un logger global con:
    - Logs en archivo (rotación diaria)
    - Logs en consola
    """

    def __init__(self, log_level="INFO", log_dir="logs", log_name="botexcel.log"):
        self.log_level = log_level.upper()
        self.log_dir = log_dir
        self.log_name = log_name

        # Crear carpeta logs si no existe
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)

        self.logger = logging.getLogger("BotExcelLogger")
        self.logger.setLevel(self.log_level)
        self.logger.propagate = False  # 🔴 IMPORTANTE: evitar duplicados

        # Evitar handlers duplicados
        if not self.logger.handlers:

            formatter = logging.Formatter(
                "%(levelname)s - %(message)s"
            )

            # ============================
            # 📄 HANDLER ARCHIVO
            # ============================
            file_path = os.path.join(self.log_dir, self.log_name)

            file_handler = TimedRotatingFileHandler(
                file_path,
                when="midnight",
                interval=1,
                backupCount=7,
                encoding="utf-8"
            )
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

            # ============================
            # 🖥️ HANDLER CONSOLA
            # ============================
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

    def get_logger(self):
        return self.logger
