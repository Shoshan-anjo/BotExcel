# gui/settings/theme_settings_view.py

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QComboBox
from qfluentwidgets import setTheme, Theme
from dotenv import load_dotenv, dotenv_values
import os


class ThemeSettingsView(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)

        load_dotenv()
        self.env_path = ".env"

        self.current_theme = os.getenv("THEME", "light").lower()

        layout = QVBoxLayout(self)
        layout.setSpacing(20)

        title = QLabel("Configuración de Tema")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Claro", "Oscuro"])
        self.theme_combo.setCurrentIndex(0 if self.current_theme == "light" else 1)
        self.theme_combo.currentIndexChanged.connect(self.change_theme)

        layout.addWidget(title)
        layout.addWidget(QLabel("Selecciona el tema de la aplicación:"))
        layout.addWidget(self.theme_combo)
        layout.addStretch()

    # -------------------------
    # Cambiar tema
    # -------------------------
    def change_theme(self, index: int):
        if index == 1:
            self.current_theme = "dark"
            setTheme(Theme.DARK)
        else:
            self.current_theme = "light"
            setTheme(Theme.LIGHT)

        self._save_theme()

    # -------------------------
    # Guardar sin borrar .env
    # -------------------------
    def _save_theme(self):
        env_vars = dotenv_values(self.env_path)
        env_vars["THEME"] = self.current_theme

        with open(self.env_path, "w", encoding="utf-8") as f:
            for k, v in env_vars.items():
                if v is not None:
                    f.write(f"{k}={v}\n")
