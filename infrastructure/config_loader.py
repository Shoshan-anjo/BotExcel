# infrastructure/config_loader.py

import os
from dotenv import load_dotenv

class ConfigLoader:
    """
    Carga variables desde .env y permite obtenerlas con config.get("KEY")
    """
    def __init__(self, env_path=".env", logger=None):
        self.logger = logger
        load_dotenv(env_path)
        self.config = os.environ

    def get(self, key, default=None):
        return self.config.get(key, default)

    def get_bool(self, key, default=False):
        val = self.config.get(key)
        if val is None:
            return default
        return str(val).lower() in ("true", "1", "yes")

    def get_int(self, key, default=0):
        val = self.config.get(key)
        try:
            return int(val)
        except:
            return default

    def get_list(self, key, separator=";"):
        value = self.config.get(key)
        if not value:
            return []
        return [x.strip() for x in value.split(separator) if x.strip()]

    def set(self, key, value):
        """Actualiza la variable en memoria y en el .env"""
        self.config[key] = str(value)
        env_path = ".env"
        lines = []
        if os.path.exists(env_path):
            with open(env_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
        found = False
        for i, line in enumerate(lines):
            if line.strip().startswith(f"{key}="):
                lines[i] = f"{key}={value}\n"
                found = True
                break
        if not found:
            lines.append(f"{key}={value}\n")
        with open(env_path, "w", encoding="utf-8") as f:
            f.writelines(lines)

# Esto asegura que ConfigLoader se pueda importar directamente
__all__ = ["ConfigLoader"]
