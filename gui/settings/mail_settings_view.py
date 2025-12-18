from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PyQt5.QtGui import QFont, QIntValidator
from PyQt5.QtCore import Qt
from qfluentwidgets import LineEdit, SwitchButton, PushButton, InfoBar, InfoBarPosition, FluentIcon
from dotenv import dotenv_values
import re
import os

EMAIL_REGEX = r"^[^@]+@[^@]+\.[^@]+$"

class UnifiedMailSettingsView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        font = QFont("Segoe UI", 10)
        font.setStyleStrategy(QFont.PreferAntialias)
        self.setFont(font)

        self.env_path = ".env"
        self.env = dotenv_values(self.env_path)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(16)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # =============================
        # Configuración de correo
        # =============================
        header_mail = QLabel("Configuración de correo")
        header_mail.setStyleSheet("font-weight: bold; font-size: 16px;")
        main_layout.addWidget(header_mail)

        # Habilitar correo
        self.enable_mail = SwitchButton("Habilitar envío de correos")
        self.enable_mail.setChecked(self.env.get("MAIL_ENABLED", "true") == "true")
        self.enable_mail.checkedChanged.connect(self._toggle_mail_inputs)
        main_layout.addWidget(self.enable_mail)

        # Usar Outlook Desktop
        self.use_outlook = SwitchButton("Usar Outlook Desktop")
        self.use_outlook.setChecked(self.env.get("USE_OUTLOOK_DESKTOP", "true") == "true")
        self.use_outlook.checkedChanged.connect(self._toggle_mail_inputs)
        main_layout.addWidget(self.use_outlook)

        # Correo remitente / usuario SMTP
        self.sender_input = LineEdit()
        self.sender_input.setPlaceholderText("Correo remitente / usuario SMTP")
        self.sender_input.setText(self.env.get("MAIL_FROM", ""))
        main_layout.addWidget(self.sender_input)

        # Contraseña SMTP (solo si no usa Outlook)
        self.password_input = LineEdit()
        self.password_input.setPlaceholderText("Contraseña SMTP")
        self.password_input.setEchoMode(LineEdit.Password)
        self.password_input.setText(self.env.get("MAIL_PASSWORD", ""))
        main_layout.addWidget(self.password_input)

        # Destinatarios
        self.recipients_input = LineEdit()
        self.recipients_input.setPlaceholderText("Destinatarios separados por coma o punto y coma")
        self.recipients_input.setText(self.env.get("MAIL_TO", ""))
        main_layout.addWidget(self.recipients_input)

        # Enviar adjunto
        self.send_attachment = SwitchButton("Enviar Excel adjunto")
        self.send_attachment.setChecked(self.env.get("MAIL_SEND_ATTACHMENT", "true") == "true")
        main_layout.addWidget(self.send_attachment)

        # Botón de prueba
        self.btn_test_email = PushButton("Probar envío")
        self.btn_test_email.setIcon(FluentIcon.MAIL)
        self.btn_test_email.setFixedHeight(36)
        self.btn_test_email.clicked.connect(self.probar_envio)
        main_layout.addWidget(self.btn_test_email)

        # =============================
        # Configuración de actualización
        # =============================
        header_refresh = QLabel("Configuración de actualización")
        header_refresh.setStyleSheet("font-weight: bold; font-size: 16px;")
        main_layout.addWidget(header_refresh)

        int_validator_repeat = QIntValidator(1, 100)
        int_validator_retry = QIntValidator(0, 1000)

        # Repeticiones
        repeat_layout = QHBoxLayout()
        repeat_layout.addWidget(QLabel("Cantidad de actualizaciones por Excel:"))
        self.refresh_repeat = LineEdit()
        self.refresh_repeat.setValidator(int_validator_repeat)
        self.refresh_repeat.setText(str(self.env.get("REFRESH_REPEAT_COUNT", "1")))
        repeat_layout.addWidget(self.refresh_repeat)
        repeat_layout.addStretch()
        main_layout.addLayout(repeat_layout)

        # Reintentos
        retry_layout = QHBoxLayout()
        retry_layout.addWidget(QLabel("Cantidad de reintentos:"))
        self.max_retries = LineEdit()
        self.max_retries.setValidator(int_validator_retry)
        self.max_retries.setText(str(self.env.get("MAX_RETRIES", "3")))
        retry_layout.addWidget(self.max_retries)

        retry_layout.addSpacing(20)
        retry_layout.addWidget(QLabel("Intervalo entre reintentos (s):"))
        self.retry_interval = LineEdit()
        self.retry_interval.setValidator(int_validator_retry)
        self.retry_interval.setText(str(self.env.get("RETRY_INTERVAL_SECONDS", "15")))
        retry_layout.addWidget(self.retry_interval)
        retry_layout.addStretch()
        main_layout.addLayout(retry_layout)

        # Validación de filas
        validate_layout = QHBoxLayout()
        self.validate_rows = SwitchButton("Validar filas después de actualizar")
        self.validate_rows.setChecked(self.env.get("VALIDATE_ROWS_AFTER_REFRESH", "true") == "true")
        self.min_rows = LineEdit()
        self.min_rows.setValidator(int_validator_repeat)
        self.min_rows.setText(str(self.env.get("MIN_ROWS_EXPECTED", "1")))
        validate_layout.addWidget(self.validate_rows)
        validate_layout.addWidget(QLabel("Mínimo de filas:"))
        validate_layout.addWidget(self.min_rows)
        validate_layout.addStretch()
        main_layout.addLayout(validate_layout)

        # Botón guardar
        self.btn_save = PushButton("Guardar configuración")
        self.btn_save.setIcon(FluentIcon.SAVE)
        self.btn_save.setFixedHeight(40)
        self.btn_save.clicked.connect(self.save_settings)
        main_layout.addWidget(self.btn_save, alignment=Qt.AlignCenter)

        main_layout.addStretch()
        self._toggle_mail_inputs()

    # -----------------------------
    # Habilitar / deshabilitar inputs según switches
    # -----------------------------
    def _toggle_mail_inputs(self):
        enabled = self.enable_mail.isChecked()
        use_outlook = self.use_outlook.isChecked()

        self.sender_input.setEnabled(enabled)
        self.recipients_input.setEnabled(enabled)
        self.password_input.setEnabled(enabled and not use_outlook)
        self.send_attachment.setEnabled(enabled and not use_outlook)
        self.btn_test_email.setEnabled(enabled)

    # -----------------------------
    # Validar correos
    # -----------------------------
    def _validar_correos(self, texto):
        correos = [c.strip() for c in re.split(r"[;,]", texto) if c.strip()]
        invalidos = [c for c in correos if not re.match(EMAIL_REGEX, c)]
        correos_unicos = list(dict.fromkeys(correos))
        return correos_unicos, invalidos

    # -----------------------------
    # Guardar configuración en .env
    # -----------------------------
    def save_settings(self):
        env = dict(self.env)

        env["MAIL_ENABLED"] = str(self.enable_mail.isChecked()).lower()
        env["USE_OUTLOOK_DESKTOP"] = str(self.use_outlook.isChecked()).lower()
        env["MAIL_FROM"] = self.sender_input.text().strip()
        env["MAIL_PASSWORD"] = self.password_input.text().strip()
        env["MAIL_TO"] = ",".join(self._validar_correos(self.recipients_input.text())[0])
        env["MAIL_SEND_ATTACHMENT"] = str(self.send_attachment.isChecked()).lower()
        env["REFRESH_REPEAT_COUNT"] = self.refresh_repeat.text() or "1"
        env["MAX_RETRIES"] = self.max_retries.text() or "3"
        env["RETRY_INTERVAL_SECONDS"] = self.retry_interval.text() or "15"
        env["VALIDATE_ROWS_AFTER_REFRESH"] = str(self.validate_rows.isChecked()).lower()
        env["MIN_ROWS_EXPECTED"] = self.min_rows.text() or "1"

        with open(self.env_path, "w", encoding="utf-8") as f:
            for k, v in env.items():
                if v is not None:
                    f.write(f"{k}={v}\n")

        InfoBar.success(title="Configuración", content="Configuración guardada correctamente", position=InfoBarPosition.TOP, parent=self)

    # -----------------------------
    # Probar envío
    # -----------------------------
    def probar_envio(self):
        from infrastructure.email_notifier import EmailNotifier
        notifier = EmailNotifier(logger=self.parent().logger, config=self.env)
        try:
            notifier.send_email(subject="Prueba de BotExcel", body="Este es un correo de prueba desde BotExcel.")
            InfoBar.success(title="Prueba", content="Correo enviado correctamente.", position=InfoBarPosition.TOP, parent=self)
        except Exception as e:
            InfoBar.error(title="Prueba fallida", content=f"Error al enviar correo: {e}", position=InfoBarPosition.TOP, parent=self)
