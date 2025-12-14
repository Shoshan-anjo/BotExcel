import pythoncom
import win32com.client as win32
import os

class EmailNotifier:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

    def send_email(self, subject, body, attachments=None):
        try:
            if not self.config.get_bool("MAIL_ENABLED", False):
                self.logger.info("MAIL_ENABLED=false → correo deshabilitado.")
                return

            to_addr = self.config.get("MAIL_TO")
            from_addr = self.config.get("MAIL_FROM")
            attach_logs = self.config.get_bool("ATTACH_LOG_ON_ERROR", False)

            if not to_addr:
                self.logger.warning("MAIL_TO vacío → no se enviará correo.")
                return

            # Preparar adjuntos
            final_attachments = []

            if attachments:
                final_attachments.extend(attachments)

            if attach_logs:
                log_dir = self.config.get("LOG_DIR", "logs")
                log_file = self.config.get("LOG_FILE", "botexcel.log")
                log_path = os.path.abspath(os.path.join(log_dir, log_file))

                if os.path.exists(log_path):
                    final_attachments.append(log_path)
                else:
                    self.logger.warning(f"No existe el log para adjuntar: {log_path}")

            pythoncom.CoInitialize()

            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.Subject = subject
            mail.Body = body
            mail.To = to_addr

            if from_addr:
                mail.SentOnBehalfOfName = from_addr

            for path in final_attachments:
                try:
                    mail.Attachments.Add(path)
                    self.logger.info(f"Adjuntando archivo: {path}")
                except Exception as e:
                    self.logger.warning(f"No se pudo adjuntar {path}: {e}")

            mail.Send()
            self.logger.info("Correo enviado correctamente.")

        except Exception as e:
            # ⚠️ NUNCA ROMPER EL BOT
            self.logger.warning(f"Fallo al enviar correo (ignorado): {e}")

        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
