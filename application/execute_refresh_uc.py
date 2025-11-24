# application/execute_refresh_uc.py

from datetime import datetime

import pytz

from domain.refresh_job import RefreshJob
from domain.exceptions import (
    RefreshJobError,
    ExcelGatewayError,
    ConfigError,
    EmailNotificationError,
)

from infrastructure.config_loader import ConfigLoader
from infrastructure.logger_service import LoggerService
from infrastructure.excel_gateway import ExcelGateway
from infrastructure.email_notifier import EmailNotifier


def execute_refresh():
    # ================================
    # 1) Cargar configuración y logger
    # ================================
    config = ConfigLoader()
    logger = LoggerService(config).get_logger()
    notifier = EmailNotifier(logger, config)

    timezone_name = config.get("TIMEZONE", "America/La_Paz")
    tz = pytz.timezone(timezone_name)

    start_global = datetime.now(tz)
    logger.info("=== INICIO BotExcel MULTI-ARCHIVO + MULTI-BACKUP ===")

    # ================================
    # 2) Obtener rutas
    # ================================
    try:
        principal_paths = config.get_excel_paths()
        backup_paths = config.get_excel_backup_paths()
    except ConfigError as e:
        logger.error(f"Error de configuración: {str(e)}")
        try:
            notifier.send_email(
                subject="BotExcel - ERROR de configuración",
                body=f"Ocurrió un error al cargar rutas de Excel:\n\n{str(e)}",
                attach_log=True,
            )
        except EmailNotificationError:
            # Si incluso el correo falla, solo logueamos
            logger.error("No se pudo enviar correo de error de configuración.")
        return

    # Validar cantidad de backups
    if backup_paths and len(principal_paths) != len(backup_paths):
        msg = (
            f"La cantidad de EXCEL_PATHS ({len(principal_paths)}) no coincide con "
            f"EXCEL_BACKUP_PATHS ({len(backup_paths)})."
        )
        logger.error(msg)
        try:
            notifier.send_email(
                subject="BotExcel - ERROR en configuración de backups",
                body=msg,
                attach_log=True,
            )
        except EmailNotificationError:
            logger.error("No se pudo enviar correo de error de backups.")
        return

    resultados = []
    dry_run = str(config.get("DRY_RUN", "false")).lower() == "true"

    # ================================
    # 3) Procesar cada archivo
    # ================================
    for idx, principal in enumerate(principal_paths):
        t_file_start = datetime.now(tz)
        backup = backup_paths[idx] if backup_paths else None

        logger.info(f"Iniciando archivo principal: {principal}")

        if dry_run:
            logger.info(f"DRY_RUN activo → se simula refresh para: {principal}")
            resultados.append(
                {
                    "archivo": principal,
                    "estado": "SIMULADO",
                    "duracion": 0,
                    "detalle": "DRY_RUN activo, no se ejecutó Excel.",
                    "fallback": False,
                }
            )
            continue

        try:
            job = RefreshJob(principal)
            gateway = ExcelGateway(logger, config)
            job.execute(gateway)

            t_file_end = datetime.now(tz)
            duracion = round((t_file_end - t_file_start).total_seconds(), 2)

            resultados.append(
                {
                    "archivo": principal,
                    "estado": "OK",
                    "duracion": duracion,
                    "fallback": False,
                }
            )

        except (ExcelGatewayError, RefreshJobError) as e:
            logger.error(f"Error con archivo principal {principal}: {str(e)}")

            # Intentar backup si existe
            if backup:
                logger.info(f"Intentando BACKUP: {backup}")
                try:
                    job_bk = RefreshJob(backup)
                    gateway = ExcelGateway(logger, config)
                    job_bk.execute(gateway)

                    t_file_end = datetime.now(tz)
                    duracion = round((t_file_end - t_file_start).total_seconds(), 2)

                    resultados.append(
                        {
                            "archivo": principal,
                            "estado": "OK (BACKUP)",
                            "duracion": duracion,
                            "fallback": True,
                            "backup_path": backup,
                        }
                    )
                except (ExcelGatewayError, RefreshJobError) as e2:
                    logger.error(f"Backup también falló para {principal}: {str(e2)}")
                    resultados.append(
                        {
                            "archivo": principal,
                            "estado": "ERROR",
                            "detalle": f"Principal: {str(e)} | Backup: {str(e2)}",
                            "fallback": True,
                        }
                    )
            else:
                resultados.append(
                    {
                        "archivo": principal,
                        "estado": "ERROR",
                        "detalle": str(e),
                        "fallback": False,
                    }
                )

    # ================================
    # 4) Resumen final
    # ================================
    end_global = datetime.now(tz)
    total_segundos = round((end_global - start_global).total_seconds(), 2)

    resumen = [
        "RESUMEN DE ACTUALIZACIÓN (MULTI-ARCHIVO + MULTI-BACKUP)",
        "",
    ]

    for r in resultados:
        if r["estado"].startswith("OK"):
            resumen.append(f"✔ {r['archivo']}")
            resumen.append(f"   Estado: {r['estado']}")
            resumen.append(f"   Duración: {r['duracion']}s")
            if r.get("fallback"):
                resumen.append(f"   Se usó backup: {r.get('backup_path')}")
            resumen.append("")
        elif r["estado"] == "SIMULADO":
            resumen.append(f"🛈 {r['archivo']}")
            resumen.append(f"   Estado: {r['estado']}")
            resumen.append(f"   Detalle: {r['detalle']}")
            resumen.append("")
        else:
            resumen.append(f"❌ {r['archivo']}")
            resumen.append("   FALLÓ")
            resumen.append(f"   Detalle: {r.get('detalle', 'Sin detalle')}")
            resumen.append(f"   Fallback usado: {r['fallback']}")
            resumen.append("")

    resumen.append(f"Tiempo total del bot: {total_segundos} segundos")
    resumen.append(f"Inicio: {start_global}")
    resumen.append(f"Fin: {end_global}")

    cuerpo = "\n".join(resumen)

    # ================================
    # 5) Enviar correo
    # ================================
    attach_log = str(config.get("ATTACH_LOG_ON_ERROR", "true")).lower() == "true"

    try:
        notifier.send_email(
            subject="BotExcel - Resumen Multi-Archivo + Multi-Backup",
            body=cuerpo,
            attach_log=attach_log,
        )
    except EmailNotificationError as e:
        logger.error(f"No se pudo enviar el correo de resumen: {str(e)}")

    logger.info("=== FIN BotExcel MULTI-ARCHIVO + MULTI-BACKUP ===")
