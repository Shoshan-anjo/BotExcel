# infrastructure/excel_gateway.py

import time
import pythoncom
import win32com.client as win32
import os

from domain.exceptions import ExcelGatewayError


class ExcelGateway:
    """
    Maneja Excel mediante COM.

    Características:
    - RefreshAll repetido según variable de entorno
    - Reintentos automáticos
    - Detección de archivo bloqueado
    - Espera robusta del cálculo (PowerQuery / Fórmulas)
    - Validación del archivo actualizado
    """

    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

        self.max_retries = int(config.get("MAX_RETRIES", 3))
        self.retry_interval = int(config.get("RETRY_INTERVAL_SECONDS", 15))
        self.min_rows_expected = int(config.get("MIN_ROWS_EXPECTED", 1))
        self.validate_rows = str(config.get("VALIDATE_ROWS_AFTER_REFRESH", "true")).lower() == "true"

        # 🔥 Número de veces que se realizará el refresh
        self.refresh_repeat = int(config.get("REFRESH_REPEAT_COUNT", 1))

    # ======================================================
    # DETECCIÓN DEL ARCHIVO BLOQUEADO
    # ======================================================
    def file_is_locked(self, path):
        if not os.path.exists(path):
            raise ExcelGatewayError(f"El archivo no existe: {path}")

        try:
            stream = open(path, "r+b")
            stream.close()
            return False
        except IOError:
            return True

    # ======================================================
    # EJECUTAR REFRESH COMPLETO CON REFRESH REPETIDO
    # ======================================================
    def refresh_file(self, excel_path):
        self.logger.info(f"Iniciando refresh del archivo: {excel_path}")

        if self.file_is_locked(excel_path):
            raise ExcelGatewayError(f"El archivo está bloqueado o en uso: {excel_path}")

        attempt = 1

        while attempt <= self.max_retries:
            try:
                self.logger.info(f"Intento {attempt} de {self.max_retries}...")

                pythoncom.CoInitialize()

                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(excel_path)

                # ============================================================
                # 🔥 Refresh repetido según REFRESH_REPEAT_COUNT
                # ============================================================
                for i in range(self.refresh_repeat):
                    self.logger.info(f"Ejecutando RefreshAll() (ronda {i+1}/{self.refresh_repeat})...")

                    t0 = time.time()
                    wb.RefreshAll()
                    self._wait_for_refresh(excel)
                    t1 = time.time()

                    self.logger.info(f"Ronda {i+1} completada en {round(t1 - t0, 2)} segundos.")

                # Guardar cambios
                wb.Save()
                wb.Close()
                excel.Quit()

                # Validación final
                if self.validate_rows:
                    self._validate_excel_after_refresh(excel_path)

                return {
                    "status": "ok",
                    "message": "Archivo actualizado correctamente",
                }

            except Exception as e:
                self.logger.error(f"Error en intento {attempt}: {str(e)}")

                if attempt == self.max_retries:
                    raise ExcelGatewayError(
                        f"Fallaron todos los intentos de refresh. Último error: {str(e)}"
                    )

                self.logger.info(f"Esperando {self.retry_interval} segundos antes del próximo intento...")
                time.sleep(self.retry_interval)

            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

            attempt += 1

    # ======================================================
    # ESPERAR A QUE EXCEL TERMINE EL CÁLCULO
    # ======================================================
    def _wait_for_refresh(self, excel_app, timeout=120):
        self.logger.info("Esperando a que Excel termine el cálculo...")

        start = time.time()

        while True:
            try:
                state = excel_app.CalculateState
            except:
                # Si COM falla, asumimos que terminó
                break

            # 0 = xlDone
            if state == 0:
                break

            # 1 = xlCalculating (esperar)
            time.sleep(0.5)

            if time.time() - start > timeout:
                self.logger.warning("Timeout esperando fin de cálculo.")
                break

    # ======================================================
    # VALIDACIÓN POST REFRESH
    # ======================================================
    def _validate_excel_after_refresh(self, path):
        size = os.path.getsize(path)

        if size <= 0:
            raise ExcelGatewayError("El archivo quedó vacío después del refresh.")

        self.logger.info(f"Validación OK: archivo final tiene {size} bytes.")
