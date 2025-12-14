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
    logger = LoggerService(config.get_log_level()).get_logger()
    notifier = EmailNotifier(logger, config)

    tz = pytz.timezone(config.get("TIMEZONE", "UTC"))
    t_start_global = datetime.now(tz)

    logger.info("=== INICIO BotExcel MULTI-ARCHIVO + MULTI-BACKUP ===")

    principals = files if files else config.get_excel_paths()
    backups = config.get_excel_backup_paths()

    if backups and len(principals) != len(backups):
        error_msg = f"La cantidad de EXCEL_PATHS ({len(principals)}) NO coincide con EXCEL_BACKUP_PATHS ({len(backups)})."
        logger.error(error_msg)
        notifier.send_email("BotExcel - ERROR en configuración", error_msg)
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
    notifier.send_email("BotExcel - Resumen Multi-Archivo + Multi-Backup", resumen)
    logger.info("=== FIN BotExcel MULTI-ARCHIVO + MULTI-BACKUP ===")
