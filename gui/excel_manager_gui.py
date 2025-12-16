from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QFileDialog
)
from PyQt5.QtGui import QFont, QRegExpValidator
from PyQt5.QtCore import QRegExp
from qfluentwidgets import (
    TableWidget, PushButton,
    LineEdit, InfoBar, InfoBarPosition,
    FluentIcon, SwitchButton
)
import json
import os

CONFIG_PATH = "config/excels.json"

class ExcelManagerView(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)

        # -------------------------
        # Fuente amigable
        # -------------------------
        font = QFont("Segoe UI", 10)
        font.setStyleStrategy(QFont.PreferAntialias)
        self.setFont(font)

        self.layout = QVBoxLayout(self)

        # -------------------------
        # Tabla
        # -------------------------
        self.table = TableWidget(self)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels([
            "Excel",
            "Backup",
            "Horario (HH:MM)",
            "Activo"
        ])
        self.layout.addWidget(self.table)

        # -------------------------
        # Botones
        # -------------------------
        btn_layout = QHBoxLayout()
        self.btn_add = PushButton("Agregar Excel", self, FluentIcon.ADD)
        self.btn_backup = PushButton("Asignar Backup", self, FluentIcon.FOLDER)
        self.btn_save = PushButton("Guardar", self, FluentIcon.SAVE)

        self.btn_add.clicked.connect(self.add_excel)
        self.btn_backup.clicked.connect(self.assign_backup)
        self.btn_save.clicked.connect(self.save)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_backup)
        btn_layout.addWidget(self.btn_save)
        self.layout.addLayout(btn_layout)

        self.load_data()

    # -------------------------
    # Cargar datos
    # -------------------------
    def load_data(self):
        self.table.setRowCount(0)
        if not os.path.exists(CONFIG_PATH):
            return

        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            excels = json.load(f).get("excels", [])

        for excel in excels:
            self.add_row(
                excel.get("path", ""),
                excel.get("backup", ""),
                excel.get("horario", ""),
                excel.get("activo", True)
            )

    # -------------------------
    # Agregar fila
    # -------------------------
    def add_row(self, path, backup, horario, activo):
        row = self.table.rowCount()
        self.table.insertRow(row)

        self.table.setItem(row, 0, self._item(path))
        self.table.setItem(row, 1, self._item(backup))

        # Horario como LineEdit con validador de HH:MM
        horario_edit = LineEdit()
        horario_edit.setPlaceholderText("HH:MM")
        horario_edit.setText(horario)
        regex = QRegExp("([01]?[0-9]|2[0-3]):[0-5][0-9]")  # HH:MM 24h
        horario_edit.setValidator(QRegExpValidator(regex))
        self.table.setCellWidget(row, 2, horario_edit)

        # Activo como SwitchButton
        toggle = SwitchButton(self)
        toggle.setChecked(activo)
        self.table.setCellWidget(row, 3, toggle)

    def _item(self, text):
        from PyQt5.QtWidgets import QTableWidgetItem
        item = QTableWidgetItem(text)
        item.setFlags(item.flags())  # no editable
        return item

    # -------------------------
    # Acciones
    # -------------------------
    def add_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Excel",
            "",
            "Excel (*.xlsx *.xlsm)"
        )
        if path:
            self.add_row(path, "", "", True)

    def assign_backup(self):
        row = self.table.currentRow()
        if row < 0:
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Backup",
            "",
            "Excel (*.xlsx *.xlsm)"
        )
        if path:
            self.table.item(row, 1).setText(path)

    # -------------------------
    # Guardar configuración
    # -------------------------
    def save(self):
        excels = []

        for row in range(self.table.rowCount()):
            horario_widget = self.table.cellWidget(row, 2)
            excels.append({
                "path": self.table.item(row, 0).text(),
                "backup": self.table.item(row, 1).text(),
                "horario": horario_widget.text() if horario_widget else "",
                "activo": self.table.cellWidget(row, 3).isChecked()
            })

        os.makedirs("config", exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"excels": excels}, f, indent=4)

        InfoBar.success(
            title="Guardado",
            content="Configuración guardada correctamente",
            position=InfoBarPosition.TOP,
            parent=self
        )
