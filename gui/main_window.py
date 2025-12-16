from PyQt5.QtWidgets import QWidget
from qfluentwidgets import (
    FluentWindow,
    FluentIcon,
    NavigationItemPosition,
    NavigationToolButton,
    setTheme,
    Theme
)
from dotenv import load_dotenv, dotenv_values
import os

from gui.excel_manager_gui import ExcelManagerView
from gui.mail_settings_view import MailSettingsView


class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()

        # -------------------------
        # Cargar variables de entorno
        # -------------------------
        load_dotenv()
        self.env_path = ".env"
        self.env = dotenv_values(self.env_path)

        # -------------------------
        # Tema
        # -------------------------
        self.current_theme = self.env.get("THEME", "light").lower()
        self._apply_theme(self.current_theme)

        # -------------------------
        # Configuración de ventana
        # -------------------------
        self.setWindowTitle("BotExcel")
        self.resize(1200, 750)

        # -------------------------
        # Vistas principales
        # -------------------------
        self.excel_view = ExcelManagerView(self)
        self.excel_view.setObjectName("excel_manager")

        self.mail_view = MailSettingsView(self)
        self.mail_view.setObjectName("mail_settings")

        # -------------------------
        # Navegación lateral
        # -------------------------
        self.addSubInterface(
            self.excel_view,
            FluentIcon.DOCUMENT,
            "Excels",
            position=NavigationItemPosition.TOP
        )

        self.addSubInterface(
            self.mail_view,
            FluentIcon.MAIL,
            "Correo",
            position=NavigationItemPosition.TOP
        )

        # -------------------------
        # Botón de cambio de tema (barra lateral abajo)
        # -------------------------
        self.theme_button = NavigationToolButton(self)
        self.theme_button.setIcon(FluentIcon.CONSTRACT)
        self.theme_button.clicked.connect(self.toggle_theme)
        self._update_theme_text()

        self.navigationInterface.addWidget(
            widget=self.theme_button,
            routeKey="theme_switch",
            position=NavigationItemPosition.BOTTOM
        )

    # -------------------------
    # Aplicar tema
    # -------------------------
    def _apply_theme(self, theme_name):
        if theme_name == "light":
            setTheme(Theme.LIGHT)
        else:
            setTheme(Theme.DARK)

    # -------------------------
    # Actualizar texto del botón de tema
    # -------------------------
    def _update_theme_text(self):
        self.theme_button.setText(f"Tema: {self.current_theme.capitalize()}")

    # -------------------------
    # Cambiar tema
    # -------------------------
    def toggle_theme(self):
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self._apply_theme(self.current_theme)
        self._update_theme_text()
        self._save_theme_env()

    # -------------------------
    # Guardar tema en .env sin eliminar otras variables
    # -------------------------
    def _save_theme_env(self):
        env_vars = dotenv_values(self.env_path)
        env_vars["THEME"] = self.current_theme
        with open(self.env_path, "w", encoding="utf-8") as f:
            for k, v in env_vars.items():
                if v is not None:
                    f.write(f"{k}={v}\n")
