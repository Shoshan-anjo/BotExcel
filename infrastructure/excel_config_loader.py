import json
import os
from domain.exceptions import ConfigError


class ExcelConfigLoader:

    def __init__(self, config_path="config/excels.json"):
        self.config_path = config_path

    def load(self):
        if not os.path.exists(self.config_path):
            raise ConfigError(f"No existe el archivo {self.config_path}")

        with open(self.config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        excels = data.get("excels", [])
        if not excels:
            raise ConfigError("No hay Excels configurados para actualizar.")

        return excels
