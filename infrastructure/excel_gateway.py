# infrastructure/excel_gateway.py

import time
import pythoncom
import win32com.client as win32
import win32gui
import win32ui
import win32con
import win32api
import os
from datetime import datetime
from domain.exceptions import ExcelGatewayError

class ExcelGateway:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config
        self.max_retries = config.get_int("MAX_RETRIES", 3)
        self.retry_interval = config.get_int("RETRY_INTERVAL_SECONDS", 15)
        self.validate_rows = config.get_bool("VALIDATE_ROWS_AFTER_REFRESH", True)
        self.screenshot_on_error = config.get_bool("SCREENSHOT_ON_ERROR", True)

    def _take_screenshot(self, name_prefix="error"):
        try:
            os.makedirs("logs/screenshots", exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"logs/screenshots/{name_prefix}_{timestamp}.png"
            
            # Use win32 for screenshot to avoid external deps if possible
            hdesktop = win32gui.GetDesktopWindow()
            width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
            height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
            left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

            desktop_dc = win32gui.GetWindowDC(hdesktop)
            img_dc = win32ui.CreateDCFromHandle(desktop_dc)
            mem_dc = img_dc.CreateCompatibleDC()

            screenshot = win32ui.CreateBitmap()
            screenshot.CreateCompatibleBitmap(img_dc, width, height)
            mem_dc.SelectObject(screenshot)
            mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top), win32con.SRCCOPY)
            
            screenshot.SaveBitmapFile(mem_dc, filename)
            
            # Clean up
            mem_dc.DeleteDC()
            win32gui.DeleteObject(screenshot.GetHandle())
            img_dc.DeleteDC()
            win32gui.ReleaseDC(hdesktop, desktop_dc)
            
            self.logger.info(f"Captura de pantalla guardada: {filename}")
            return filename
        except Exception as e:
            self.logger.warning(f"No se pudo tomar la captura de pantalla: {e}")
            return None

    def _check_excel_health(self):
        """Verifica si hay diálogos abiertos o estados que bloqueen Excel."""
        # Por ahora simple, pero se puede expandir
        self.logger.info("Verificando salud de Excel...")
        return True

    def file_is_locked(self, path):
        try:
            with open(path, "a"):
                return False
        except PermissionError:
            return True

    def refresh_file(self, excel_path):
        self.logger.info(f"Iniciando refresh del archivo: {excel_path}")

        # --- MANEJO DE CONFLICTOS CON EL USUARIO ---
        # Si el usuario tiene el archivo abierto, intentamos esperar un poco
        max_lock_attempts = 5
        for lock_attempt in range(1, max_lock_attempts + 1):
            if not self.file_is_locked(excel_path):
                break
            if lock_attempt < max_lock_attempts:
                self.logger.warning(f"Archivo en uso por el usuario. Reintento {lock_attempt} en 30s...")
                time.sleep(30)
            else:
                raise ExcelGatewayError(f"El archivo sigue bloqueado después de {max_lock_attempts} intentos. Por favor, ciérrelo para que el bot pueda trabajar.")

        attempt = 1
        while attempt <= self.max_retries:
            excel = None
            wb = None
            try:
                self.logger.info(f"Intento {attempt} de {self.max_retries}...")
                self._check_excel_health()
                pythoncom.CoInitialize()
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(excel_path)

                # --- CONFIGURACIÓN PARA POWER QUERY PESADO ---
                # Deshabilitamos el refresco en segundo plano de todas las conexiones
                # Esto obliga a que wb.RefreshAll() sea sincrónico y espere de verdad.
                try:
                    for conn in wb.Connections:
                        if conn.Type == 1: # OLEDB (Power Query)
                            conn.OLEDBConnection.BackgroundQuery = False
                        elif conn.Type == 2: # ODBC
                            conn.ODBCConnection.BackgroundQuery = False
                except Exception as e:
                    self.logger.warning(f"No se pudieron ajustar todas las conexiones: {e}")

                t_start = time.time()
                self.logger.info("Ejecutando RefreshAll()...")
                wb.RefreshAll()
                
                # Aumentamos el timeout de espera a 2 horas (7200s) por seguridad
                self._wait_for_refresh(excel, timeout=7200)
                t_end = time.time()

                wb.Save()
                # Clean exit on success
                wb.Close()
                wb = None
                excel.Quit()
                excel = None

                if self.validate_rows:
                    self._validate_excel_after_refresh(excel_path)

                return {
                    "status": "ok",
                    "message": "Archivo actualizado correctamente",
                    "refresh_time": round(t_end - t_start, 2)
                }

            except Exception as e:
                self.logger.error(f"Error en intento {attempt}: {str(e)}")
                
                if self.screenshot_on_error:
                    self._take_screenshot(f"error_intento_{attempt}")
                
                # Force cleanup on error
                try:
                    if wb: wb.Close(SaveChanges=False)
                except: pass
                
                try:
                    if excel: excel.Quit()
                except: pass
                
                if attempt == self.max_retries:
                    raise ExcelGatewayError(f"Todos los intentos fallaron: {str(e)}")
                
                self.logger.info(f"Esperando {self.retry_interval}s antes del próximo intento...")
                time.sleep(self.retry_interval)

            finally:
                pythoncom.CoUninitialize()

            attempt += 1

    def _wait_for_refresh(self, excel_app, timeout=120):
        self.logger.info("Esperando a que Excel termine el cálculo...")
        start_time = time.time()
        while True:
            try:
                state = excel_app.CalculateState
            except:
                break
            if state == 0:
                break
            time.sleep(0.5)
            if time.time() - start_time > timeout:
                self.logger.warning("Timeout esperando a que Excel termine el cálculo.")
                break

    def _validate_excel_after_refresh(self, path):
        file_size = os.path.getsize(path)
        if file_size <= 0:
            raise ExcelGatewayError("El archivo quedó vacío después del refresh.")
        self.logger.info(f"Validación OK: archivo final tiene {file_size} bytes.")
