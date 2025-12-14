# infrastructure/scheduler_service.py

import time
import threading
from datetime import datetime
import schedule
from application.scheduler_uc import SchedulerUseCase
from application.execute_refresh_uc import execute_refresh
from infrastructure.logger_service import LoggerService
from infrastructure.config_loader import ConfigLoader

class SchedulerService:
    """
    Servicio que ejecuta los Excel según los horarios configurados
    en schedule.json (SchedulerUseCase).
    """

    def __init__(self, config=None, logger=None):
        self.config = config or ConfigLoader()
        self.logger = logger or LoggerService(self.config.get("LOG_LEVEL", "INFO")).get_logger()
        self.scheduler_uc = SchedulerUseCase()
        self.jobs_threads = []

    # --------------------------
    # Registrar jobs en schedule
    # --------------------------
    def _register_jobs(self):
        active_jobs = self.scheduler_uc.get_active_jobs()

        if not active_jobs:
            self.logger.warning("No hay jobs activos configurados en schedule.json")
            return

        for job in active_jobs:
            excel_path = job["excel_path"]
            horario = job["horario"]
            backup_path = job.get("backup_path")

            self.logger.info(f"Programando job: {excel_path} a las {horario}")

            # schedule.every().day.at(horario) ejecuta execute_refresh
            schedule.every().day.at(horario).do(self._run_job_threaded, excel_path=excel_path, backup_path=backup_path)

    # --------------------------
    # Ejecutar job en hilo separado
    # --------------------------
    def _run_job_threaded(self, excel_path, backup_path=None):
        def task():
            self.logger.info(f"=== INICIO JOB PROGRAMADO: {excel_path} ===")
            try:
                execute_refresh()  # Aquí se puede pasar parámetros si quieres actualizar solo este archivo
            except Exception as e:
                self.logger.error(f"Error ejecutando job {excel_path}: {str(e)}")
            self.logger.info(f"=== FIN JOB PROGRAMADO: {excel_path} ===")

        t = threading.Thread(target=task)
        t.start()
        self.jobs_threads.append(t)

    # --------------------------
    # Iniciar el scheduler
    # --------------------------
    def start(self):
        self.logger.info("Scheduler ACTIVADO")
        self._register_jobs()

        while True:
            schedule.run_pending()
            time.sleep(1)

    # --------------------------
    # Método para ejecutar desde GUI (no bloqueante)
    # --------------------------
    def start_in_thread(self):
        t = threading.Thread(target=self.start, daemon=True)
        t.start()
        return t
