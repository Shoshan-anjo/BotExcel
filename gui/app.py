import sys
import ctypes
from PyQt5.QtWidgets import QApplication
from qfluentwidgets import setTheme, Theme

from gui.main_window import MainWindow


def run():
    # Establecer ID de aplicación para que Windows reconozca "Pivoty" en las notificaciones
    myappid = 'shohan.pivoty.bot.1.0' 
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication(sys.argv)

    # Tema inicial (oscuro por defecto)
    setTheme(Theme.DARK)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    run()
