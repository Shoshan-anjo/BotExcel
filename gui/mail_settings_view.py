from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout
from PyQt5.QtGui import QFont, QIntValidator
from qfluentwidgets import (
    LineEdit, SwitchButton, PushButton, InfoBar, 
    InfoBarPosition, FluentIcon, TitleLabel, SubtitleLabel, BodyLabel
)
from dotenv import dotenv_values
from PyQt5.QtCore import Qt
from infrastructure.startup_manager import StartupManager

class MailSettingsView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        # -------------------------
        # Fuente amigable para toda la vista
        # -------------------------
        font = QFont("Segoe UI", 10)
        font.setStyleStrategy(QFont.PreferAntialias)
        self.setFont(font)

        self.env_path = ".env"
        self.env = dotenv_values(self.env_path)
        self.startup_manager = StartupManager()

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # -------------------------
        # Encabezado Correo
        # -------------------------
        header_mail = TitleLabel("Configuración de Notificaciones")
        main_layout.addWidget(header_mail)

        subtitle_mail = SubtitleLabel("Define quién recibe los reportes. Se usa Outlook automáticamente.")
        subtitle_mail.setTextColor("#808080", "#a0a0a0")
        main_layout.addWidget(subtitle_mail)
        main_layout.addSpacing(5)

        # Destinatarios
        self.recipients_input = LineEdit()
        self.recipients_input.setPlaceholderText("Destinatarios separados por coma")
        self.recipients_input.setText(self.env.get("MAIL_TO", ""))
        
        # Layout para destinatarios + Botón de Prueba
        to_layout = QHBoxLayout()
        to_layout.addWidget(self.recipients_input)
        
        self.btn_test = PushButton("Probar Envío")
        self.btn_test.setIcon(FluentIcon.SEND)
        self.btn_test.setFixedWidth(120)
        self.btn_test.clicked.connect(self.send_test_email)
        to_layout.addWidget(self.btn_test)
        
        main_layout.addLayout(to_layout)

        # -------------------------
        # Toggles de Configuración
        # -------------------------
        main_layout.addSpacing(5)
        
        # Grupo 1: General
        group_general = QHBoxLayout()
        self.enable_mail = SwitchButton(text="Habilitar Notificaciones")
        self.enable_mail.setChecked(self.env.get("MAIL_ENABLED", "true") == "true")
        self.run_at_startup = SwitchButton(text="Iniciar con Windows")
        self.run_at_startup.setChecked(self.startup_manager.is_startup_enabled())
        group_general.addWidget(self.enable_mail)
        group_general.addWidget(self.run_at_startup)
        group_general.addStretch()
        main_layout.addLayout(group_general)

        # Grupo 2: Contenido del Reporte
        main_layout.addWidget(BodyLabel("Contenido del Reporte:"))
        content_layout = QHBoxLayout()
        self.send_attachment = SwitchButton(text="Adjuntar Excel")
        self.send_attachment.setChecked(self.env.get("MAIL_SEND_ATTACHMENT", "true") == "true")
        
        self.include_screenshots = SwitchButton(text="Capturas de Error")
        self.include_screenshots.setChecked(self.env.get("MAIL_INCLUDE_SCREENSHOTS", "true") == "true")
        
        self.include_logs = SwitchButton(text="Snippet de Logs")
        self.include_logs.setChecked(self.env.get("MAIL_INCLUDE_LOG_SNIPPET", "true") == "true")

        content_layout.addWidget(self.send_attachment)
        content_layout.addWidget(self.include_screenshots)
        content_layout.addWidget(self.include_logs)
        content_layout.addStretch()
        main_layout.addLayout(content_layout)

        # -------------------------
        # Estado de Outlook
        # -------------------------
        status_layout = QHBoxLayout()
        self.status_label = BodyLabel("Estado de Outlook: Verificando...")
        status_layout.addWidget(self.status_label)
        main_layout.addLayout(status_layout)
        self.check_outlook_status()

        # -------------------------
        # Encabezado Actualización
        # -------------------------
        main_layout.addSpacing(20)
        header_refresh = TitleLabel("Parámetros de Robustez")
        main_layout.addWidget(header_refresh)

        subtitle_refresh = SubtitleLabel("Configura reintentos y validaciones automáticas.")
        subtitle_refresh.setTextColor("#808080", "#a0a0a0")
        main_layout.addWidget(subtitle_refresh)
        main_layout.addSpacing(5)

        int_validator_repeat = QIntValidator(1, 100)
        int_validator_retry = QIntValidator(0, 1000)

        # Filas de configuración
        repeat_layout = QHBoxLayout()
        repeat_label = BodyLabel("Cant. Actualizaciones:")
        self.refresh_repeat = LineEdit()
        self.refresh_repeat.setFixedWidth(100)
        self.refresh_repeat.setText(str(self.env.get("REFRESH_REPEAT_COUNT", 1)))
        self.refresh_repeat.setValidator(int_validator_repeat)
        repeat_layout.addWidget(repeat_label)
        repeat_layout.addWidget(self.refresh_repeat)
        
        retry_label = BodyLabel("Reintentos:")
        self.max_retries = LineEdit()
        self.max_retries.setFixedWidth(100)
        self.max_retries.setText(str(self.env.get("MAX_RETRIES", 3)))
        self.max_retries.setValidator(int_validator_retry)
        repeat_layout.addSpacing(20)
        repeat_layout.addWidget(retry_label)
        repeat_layout.addWidget(self.max_retries)
        
        interval_label = BodyLabel("Intervalo (s):")
        self.retry_interval = LineEdit()
        self.retry_interval.setFixedWidth(100)
        self.retry_interval.setText(str(self.env.get("RETRY_INTERVAL_SECONDS", 15)))
        self.retry_interval.setValidator(int_validator_retry)
        repeat_layout.addSpacing(20)
        repeat_layout.addWidget(interval_label)
        repeat_layout.addWidget(self.retry_interval)
        repeat_layout.addStretch()
        main_layout.addLayout(repeat_layout)

        # Validación de filas
        main_layout.addSpacing(10)
        validate_layout = QHBoxLayout()
        self.validate_rows = SwitchButton(text="Validar filas mínimas")
        self.validate_rows.setChecked(self.env.get("VALIDATE_ROWS_AFTER_REFRESH", "true") == "true")
        self.min_rows = LineEdit()
        self.min_rows.setPlaceholderText("1")
        self.min_rows.setFixedWidth(80)
        self.min_rows.setText(str(self.env.get("MIN_ROWS_EXPECTED", 1)))
        self.min_rows.setValidator(int_validator_repeat)

        validate_layout.addWidget(self.validate_rows)
        validate_layout.addSpacing(10)
        validate_layout.addWidget(BodyLabel("Mínimo:"))
        validate_layout.addWidget(self.min_rows)
        validate_layout.addStretch()
        main_layout.addLayout(validate_layout)

        # -------------------------
        # Botón Guardar
        # -------------------------
        main_layout.addSpacing(30)
        self.btn_save = PushButton("Guardar toda la configuración")
        self.btn_save.setIcon(FluentIcon.SAVE)
        self.btn_save.setFixedHeight(45)
        self.btn_save.setFixedWidth(300)
        main_layout.addWidget(self.btn_save, alignment=Qt.AlignCenter)
        self.btn_save.clicked.connect(self.save_settings)
        main_layout.addStretch()

    # -------------------------
    # Métodos Auxiliares
    # -------------------------
    def check_outlook_status(self):
        """Verifica si Outlook está disponible."""
        try:
            import win32com.client
            # Solo intentamos crear el dispatch para ver si responde
            win32com.client.GetActiveObject("Outlook.Application")
            self.status_label.setText("Estado de Outlook: Conectado (Activo) ✅")
            self.status_label.setTextColor("#2ecc71", "#2ecc71")
        except:
            try:
                import win32com.client
                win32com.client.Dispatch("Outlook.Application")
                self.status_label.setText("Estado de Outlook: Disponible (Bot Enviará ok) 🆗")
                self.status_label.setTextColor("#3498db", "#3498db")
            except:
                self.status_label.setText("Estado de Outlook: No detectado o configurado ❌")
                self.status_label.setTextColor("#e74c3c", "#e74c3c")

    def send_test_email(self):
        """Envía un correo de prueba."""
        try:
            from infrastructure.email_notifier import EmailNotifier
            from infrastructure.config_loader import ConfigLoader
            from infrastructure.logger_service import LoggerService
            
            # Guardamos temporalmente para que el notifier lea los datos actuales
            self.save_settings(silent=True)
            
            config = ConfigLoader()
            logger = LoggerService().get_logger()
            notifier = EmailNotifier(logger, config)
            
            notifier.send_email(
                "Pivoty - Prueba de Notificación",
                "Este es un correo de prueba enviado desde la configuración de Pivoty.\n"
                "Para que el envío funcione 24/7 sin ventanas abiertas, asegúrate de tener una cuenta de Outlook configurada en este equipo."
            )
            
            InfoBar.success(
                title="Prueba enviada",
                content="Se ha enviado el correo de prueba. Revisa tu buzón.",
                position=InfoBarPosition.TOP,
                parent=self
            )
        except Exception as e:
            InfoBar.error(
                title="Error de prueba",
                content=f"No se pudo enviar: {str(e)}",
                position=InfoBarPosition.TOP,
                parent=self
            )

    # -------------------------
    # Guardar configuración
    # -------------------------
    def save_settings(self, silent=False):
        env = dict(self.env)

        env["MAIL_TO"] = self.recipients_input.text()
        env["MAIL_ENABLED"] = str(self.enable_mail.isChecked()).lower()
        env["USE_OUTLOOK_DESKTOP"] = "true" # Siempre usar Outlook para simplicidad
        env["MAIL_SEND_ATTACHMENT"] = str(self.send_attachment.isChecked()).lower()
        env["MAIL_INCLUDE_SCREENSHOTS"] = str(self.include_screenshots.isChecked()).lower()
        env["MAIL_INCLUDE_LOG_SNIPPET"] = str(self.include_logs.isChecked()).lower()
        
        env["REFRESH_REPEAT_COUNT"] = self.refresh_repeat.text() or "1"
        env["MAX_RETRIES"] = self.max_retries.text() or "3"
        env["RETRY_INTERVAL_SECONDS"] = self.retry_interval.text() or "15"
        env["VALIDATE_ROWS_AFTER_REFRESH"] = str(self.validate_rows.isChecked()).lower()
        env["MIN_ROWS_EXPECTED"] = self.min_rows.text() or "1"

        # Configurar inicio automático
        self.startup_manager.set_startup(self.run_at_startup.isChecked())

        with open(self.env_path, "w", encoding="utf-8") as f:
            for k, v in env.items():
                if v is not None:
                    f.write(f"{k}={v}\n")

        if not silent:
            InfoBar.success(
                title="Configuración",
                content="Configuración guardada correctamente",
                position=InfoBarPosition.TOP,
                parent=self
            )
