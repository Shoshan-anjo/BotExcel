# application/execute_refresh_uc.py

from datetime import datetime
import pytz
from domain.refresh_job import RefreshJob
from infrastructure.config_loader import ConfigLoader
from infrastructure.logger_service import LoggerService
from infrastructure.email_notifier import EmailNotifier
from infrastructure.excel_gateway import ExcelGateway

def execute_refresh(files=None):
    config = ConfigLoader()
    logger = LoggerService(config.get("LOG_LEVEL", "INFO")).get_logger()
    notifier = EmailNotifier(logger, config)

    tz = pytz.timezone(config.get("TIMEZONE", "UTC"))
    t_start_global = datetime.now(tz)

    from application.scheduler_uc import SchedulerUseCase
    scheduler_uc = SchedulerUseCase()
    all_jobs = scheduler_uc._load_jobs()
    
    # Solo los activos si no se especifican archivos
    principals = files if files else [j["path"] for j in all_jobs if j.get("activo", True)]
    backups = [j.get("backup", "") for j in all_jobs if j.get("activo", True)] if not files else []

    if not principals:
        logger.info("No hay archivos activos para actualizar.")
        return

    resultados = []

    for idx, principal_path in enumerate(principals):
        backup_path = backups[idx] if backups else None
        t_file_start = datetime.now(tz)
        logger.info(f"Iniciando archivo principal: {principal_path}")

        try:
            job = RefreshJob(principal_path)
            gateway = ExcelGateway(logger, config)
            result_principal = job.execute(gateway)

            t_file_end = datetime.now(tz)
            resultados.append({
                "archivo": principal_path,
                "estado": "OK",
                "duracion": round((t_file_end - t_file_start).total_seconds(), 2),
                "refresh_time": result_principal["refresh_time"],
                "fallback": False
            })

        except Exception as e:
            logger.error(f"Error con archivo principal {principal_path}: {str(e)}")
            if backup_path:
                logger.info(f"Intentando BACKUP: {backup_path}")
                try:
                    job_bk = RefreshJob(backup_path)
                    gateway = ExcelGateway(logger, config)
                    result_backup = job_bk.execute(gateway)
                    t_file_end = datetime.now(tz)
                    resultados.append({
                        "archivo": principal_path,
                        "estado": "OK (BACKUP)",
                        "duracion": round((t_file_end - t_file_start).total_seconds(), 2),
                        "refresh_time": result_backup["refresh_time"],
                        "fallback": True,
                        "backup_path": backup_path
                    })
                except Exception as e2:
                    logger.error(f"Backup también falló: {str(e2)}")
                    resultados.append({
                        "archivo": principal_path,
                        "estado": "ERROR",
                        "error": f"Principal: {str(e)} | Backup: {str(e2)}",
                        "fallback": True
                    })
            else:
                resultados.append({
                    "archivo": principal_path,
                    "estado": "ERROR",
                    "error": str(e),
                    "fallback": False
                })

    # Resumen final
    t_end_global = datetime.now(tz)
    total_global = round((t_end_global - t_start_global).total_seconds(), 2)
    resumen = "RESUMEN DE ACTUALIZACIÓN\n\n"
    for r in resultados:
        if r["estado"].startswith("OK"):
            resumen += f"✔ {r['archivo']}\n   Estado: {r['estado']}\n   Duración: {r['duracion']}s\n   Tiempo Refresh: {r['refresh_time']}s\n"
            if r.get("fallback"):
                resumen += f"   Se usó BACKUP: {r['backup_path']}\n"
            resumen += "\n"
        else:
            resumen += f"❌ {r['archivo']}\n   FALLÓ\n   Detalle: {r['error']}\n   Fallback usado: {r['fallback']}\n\n"

    resumen += f"Tiempo total del bot: {total_global} segundos\nInicio: {t_start_global}\nFin: {t_end_global}\n"
    
    # Recolectar archivos exitosos para adjuntar (si está habilitado en config)
    excel_attachments = []
    if config.get_bool("MAIL_SEND_ATTACHMENT", True):
        for r in resultados:
            if r["estado"].startswith("OK"):
                excel_attachments.append(r["archivo"])

    notifier.send_email("Pivoty - Resumen Multi-Archivo + Multi-Backup", resumen, attachments=excel_attachments)
    logger.info("=== FIN Pivoty MULTI-ARCHIVO + MULTI-BACKUP ===")
