# infrastructure/email_notifier.py

import os
import pythoncom
import win32com.client as win32
from domain.exceptions import EmailNotificationError


class EmailNotifier:
    """
    Envío de correos vía Outlook Desktop.

    Características:
    - SIN uso de SentOnBehalfOfName (evita errores SEND AS DENIED)
    - Usa la cuenta predeterminada de Outlook (100% compatible)
    - Adjunta logs solo si existen
    - Inicialización COM robusta
    """

    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

    def send_email(self, subject, body, attach_log=False):
        try:
            # -----------------------------------------------
            # Validar si el envío está habilitado
            # -----------------------------------------------
            if str(self.config.get("MAIL_ENABLED", "false")).lower() != "true":
                self.logger.info("MAIL_ENABLED = false → No se enviará correo.")
                return

            recipient = self.config.get("MAIL_TO")
            if not recipient:
                raise EmailNotificationError(
                    "MAIL_TO no está definido en .env → No se puede enviar correo."
                )

            # -----------------------------------------------
            # Inicializar COM
            # -----------------------------------------------
            try:
                pythoncom.CoInitialize()
            except:
                # Outlook puede ya estar inicializado
                pass

            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            # -----------------------------------------------
            # Construir el correo (sin forcings de remitente)
            # -----------------------------------------------
            mail.Subject = subject
            mail.Body = body
            mail.To = recipient

            # Importante:
            # Outlook enviará desde la cuenta activa del usuario.
            # NO usar .Sender ni .SentOnBehalfOfName:
            #
            # mail.Sender = ...
            # mail.SentOnBehalfOfName = ...
            #
            # Esto causa SEND-AS DENIED si no hay permisos.

            # -----------------------------------------------
            # Adjuntar log si corresponde o otro archivo
            # -----------------------------------------------

            #if attach_log:
                #log_dir = self.config.get("LOG_DIR", "logs")
                #log_file = self.config.get("LOG_FILE", "botexcel.log")
                #abs_path = os.path.abspath(os.path.join(log_dir, log_file))

                #if os.path.exists(abs_path):
                #    try:
                #        self.logger.info(f"Adjuntando log: {abs_path}")
                #        mail.Attachments.Add(abs_path)
                #    except Exception as e:
                #        self.logger.warning(f"No se pudo adjuntar el log: {str(e)}")
                #else:
                #    self.logger.warning(f"No existe log para adjuntar: {abs_path}")

            # -----------------------------------------------
            # Enviar
            # -----------------------------------------------
            mail.Send()
            self.logger.info("Correo enviado correctamente mediante Outlook Desktop.")

        except Exception as e:
            raise EmailNotificationError(f"Error enviando correo: {str(e)}")

        finally:
            # -----------------------------------------------
            # Liberar COM
            # -----------------------------------------------
            try:
                pythoncom.CoUninitialize()
            except:
                pass
