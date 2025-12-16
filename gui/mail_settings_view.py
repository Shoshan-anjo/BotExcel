from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PyQt5.QtGui import QFont, QIntValidator
from qfluentwidgets import LineEdit, SwitchButton, PushButton, InfoBar, InfoBarPosition, FluentIcon
from dotenv import dotenv_values
from PyQt5.QtCore import Qt

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

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # -------------------------
        # Encabezado Correo
        # -------------------------
        header_mail = QLabel("Configuración de Correo")
        header_mail.setStyleSheet("font-weight: bold; font-size: 16px;")
        main_layout.addWidget(header_mail)

        # Correo remitente
        self.sender_input = LineEdit()
        self.sender_input.setPlaceholderText("Correo remitente (Outlook)")
        self.sender_input.setText(self.env.get("MAIL_FROM", ""))
        main_layout.addWidget(self.sender_input)

        # Destinatarios
        self.recipients_input = LineEdit()
        self.recipients_input.setPlaceholderText("Destinatarios separados por coma")
        self.recipients_input.setText(self.env.get("MAIL_TO", ""))
        main_layout.addWidget(self.recipients_input)

        # SwitchButton
        switch_layout = QHBoxLayout()
        switch_layout.setSpacing(15)
        self.enable_mail = SwitchButton(text="Enviar correos")
        self.enable_mail.setChecked(self.env.get("MAIL_ENABLED", "true") == "true")
        self.send_attachment = SwitchButton(text="Enviar Excel adjunto")
        self.send_attachment.setChecked(self.env.get("MAIL_SEND_ATTACHMENT", "true") == "true")
        switch_layout.addWidget(self.enable_mail)
        switch_layout.addWidget(self.send_attachment)
        switch_layout.addStretch()
        main_layout.addLayout(switch_layout)

        # -------------------------
        # Encabezado Actualización
        # -------------------------
        header_refresh = QLabel("Configuración de actualización")
        header_refresh.setStyleSheet("font-weight: bold; font-size: 16px;")
        main_layout.addWidget(header_refresh)

        int_validator_repeat = QIntValidator(1, 100)
        int_validator_retry = QIntValidator(0, 1000)

        # Cantidad de actualizaciones por Excel
        repeat_layout = QHBoxLayout()
        repeat_label = QLabel("Cantidad de actualizaciones por Excel:")
        self.refresh_repeat = LineEdit()
        self.refresh_repeat.setPlaceholderText("1")
        self.refresh_repeat.setText(str(self.env.get("REFRESH_REPEAT_COUNT", 1)))
        self.refresh_repeat.setValidator(int_validator_repeat)
        repeat_layout.addWidget(repeat_label)
        repeat_layout.addWidget(self.refresh_repeat)
        repeat_layout.addStretch()
        main_layout.addLayout(repeat_layout)

        # Reintentos y intervalo
        retry_layout = QHBoxLayout()
        retry_label = QLabel("Cantidad de reintentos:")
        self.max_retries = LineEdit()
        self.max_retries.setPlaceholderText("3")
        self.max_retries.setText(str(self.env.get("MAX_RETRIES", 3)))
        self.max_retries.setValidator(int_validator_retry)

        interval_label = QLabel("Intervalo entre reintentos (s):")
        self.retry_interval = LineEdit()
        self.retry_interval.setPlaceholderText("15")
        self.retry_interval.setText(str(self.env.get("RETRY_INTERVAL_SECONDS", 15)))
        self.retry_interval.setValidator(int_validator_retry)

        retry_layout.addWidget(retry_label)
        retry_layout.addWidget(self.max_retries)
        retry_layout.addSpacing(20)
        retry_layout.addWidget(interval_label)
        retry_layout.addWidget(self.retry_interval)
        retry_layout.addStretch()
        main_layout.addLayout(retry_layout)

        # Validación de filas
        validate_layout = QHBoxLayout()
        self.validate_rows = SwitchButton(text="Validar filas después de actualizar")
        self.validate_rows.setChecked(self.env.get("VALIDATE_ROWS_AFTER_REFRESH", "true") == "true")
        self.min_rows = LineEdit()
        self.min_rows.setPlaceholderText("1")
        self.min_rows.setText(str(self.env.get("MIN_ROWS_EXPECTED", 1)))
        self.min_rows.setValidator(int_validator_repeat)

        validate_layout.addWidget(self.validate_rows)
        validate_layout.addSpacing(10)
        validate_layout.addWidget(QLabel("Mínimo de filas:"))
        validate_layout.addWidget(self.min_rows)
        validate_layout.addStretch()
        main_layout.addLayout(validate_layout)

        # -------------------------
        # Botón Guardar
        # -------------------------
        self.btn_save = PushButton("Guardar configuración")
        self.btn_save.setIcon(FluentIcon.SAVE)
        self.btn_save.setFixedHeight(40)
        main_layout.addWidget(self.btn_save, alignment=Qt.AlignCenter)
        self.btn_save.clicked.connect(self.save_settings)
        main_layout.addStretch()

    # -------------------------
    # Guardar configuración
    # -------------------------
    def save_settings(self):
        env = dict(self.env)

        env["MAIL_FROM"] = self.sender_input.text()
        env["MAIL_TO"] = self.recipients_input.text()
        env["MAIL_ENABLED"] = str(self.enable_mail.isChecked()).lower()
        env["MAIL_SEND_ATTACHMENT"] = str(self.send_attachment.isChecked()).lower()
        env["REFRESH_REPEAT_COUNT"] = self.refresh_repeat.text() or "1"
        env["MAX_RETRIES"] = self.max_retries.text() or "3"
        env["RETRY_INTERVAL_SECONDS"] = self.retry_interval.text() or "15"
        env["VALIDATE_ROWS_AFTER_REFRESH"] = str(self.validate_rows.isChecked()).lower()
        env["MIN_ROWS_EXPECTED"] = self.min_rows.text() or "1"

        with open(".env", "w", encoding="utf-8") as f:
            for k, v in env.items():
                if v is not None:
                    f.write(f"{k}={v}\n")

        InfoBar.success(
            title="Configuración",
            content="Configuración guardada correctamente",
            position=InfoBarPosition.TOP,
            parent=self
        )
