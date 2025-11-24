# infrastructure/config_loader.py

import os
from dotenv import load_dotenv
from domain.exceptions import ConfigError


class ConfigLoader:
    """
    Carga y valida las variables de entorno desde .env
    """

    def __init__(self, env_path=".env"):
        if not os.path.exists(env_path):
            raise ConfigError(f"No se encontró el archivo .env en: {env_path}")

        load_dotenv(env_path)

        # Diccionario interno con TODAS las variables
        self.config = {
            "EXCEL_PATH": os.getenv("EXCEL_PATH"),
            "EXCEL_PATHS": os.getenv("EXCEL_PATHS"),
            "EXCEL_BACKUP_PATH": os.getenv("EXCEL_BACKUP_PATH"),
            "EXCEL_BACKUP_PATHS": os.getenv("EXCEL_BACKUP_PATHS"),

            "MAIL_ENABLED": os.getenv("MAIL_ENABLED", "false"),
            "MAIL_FROM": os.getenv("MAIL_FROM"),
            "MAIL_TO": os.getenv("MAIL_TO"),
            "USE_OUTLOOK_DESKTOP": os.getenv("USE_OUTLOOK_DESKTOP", "true"),

            "LOG_LEVEL": os.getenv("LOG_LEVEL", "INFO"),
            "LOG_DIR": os.getenv("LOG_DIR", "logs"),
            "LOG_FILE": os.getenv("LOG_FILE", "botexcel.log"),
            "ATTACH_LOG_ON_ERROR": os.getenv("ATTACH_LOG_ON_ERROR", "true"),

            "MAX_RETRIES": os.getenv("MAX_RETRIES", "3"),
            "RETRY_INTERVAL_SECONDS": os.getenv("RETRY_INTERVAL_SECONDS", "15"),
            "VALIDATE_ROWS_AFTER_REFRESH": os.getenv("VALIDATE_ROWS_AFTER_REFRESH", "true"),
            "MIN_ROWS_EXPECTED": os.getenv("MIN_ROWS_EXPECTED", "1"),

            "TIMEZONE": os.getenv("TIMEZONE", "America/La_Paz"),
            "DRY_RUN": os.getenv("DRY_RUN", "false"),

            # ⚡ NUEVO → Cantidad de repeticiones de RefreshAll()
            "REFRESH_REPEAT_COUNT": os.getenv("REFRESH_REPEAT_COUNT", "1"),
        }

    # ============================================================
    def get_excel_paths(self):
        raw = self.config["EXCEL_PATHS"]
        if not raw:
            raise ConfigError("No se definió EXCEL_PATHS en el archivo .env")

        paths = [p.strip() for p in raw.split(";") if p.strip()]
        if not paths:
            raise ConfigError("EXCEL_PATHS está vacío o mal configurado")

        return paths

    # ============================================================
    def get_excel_backup_paths(self):
        raw = self.config["EXCEL_BACKUP_PATHS"]
        if not raw:
            return []

        return [p.strip() for p in raw.split(";") if p.strip()]

    # ============================================================
    def is_mail_enabled(self):
        return self.config["MAIL_ENABLED"].lower() == "true"

    def get_mail_from(self):
        return self.config["MAIL_FROM"]

    def get_mail_to(self):
        return self.config["MAIL_TO"]

    def get_log_level(self):
        return self.config["LOG_LEVEL"]

    def use_outlook_desktop(self):
        return self.config["USE_OUTLOOK_DESKTOP"].lower() == "true"

    # ============================================================
    # MÉTODO GENÉRICO
    def get(self, key, default=None):
        return self.config.get(key, default)
