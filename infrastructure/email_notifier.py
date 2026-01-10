# infrastructure/email_notifier.py
import os
import smtplib
from email.message import EmailMessage
import pythoncom
import win32com.client as win32

class EmailNotifier:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

    def _get_log_snippet(self, lines=20):
        try:
            log_dir = self.config.get("LOG_DIR", "logs")
            log_file = self.config.get("LOG_FILE", "botexcel.log")
            log_path = os.path.abspath(os.path.join(log_dir, log_file))
            if os.path.exists(log_path):
                with open(log_path, "r", encoding="utf-8") as f:
                    content = f.readlines()
                    return "".join(content[-lines:])
        except:
            pass
        return "No se pudo recuperar el fragmento del log."

    def send_email(self, subject, body, attachments=None):
        mail_enabled = self.config.get_bool("MAIL_ENABLED", False)
        use_outlook = self.config.get_bool("USE_OUTLOOK_DESKTOP", True)

        if not mail_enabled:
            self.logger.info("MAIL_ENABLED=false → correo deshabilitado.")
            return

        to_addrs = [addr.strip() for addr in self.config.get("MAIL_TO", "").split(",") if addr.strip()]
        from_addr = self.config.get("MAIL_FROM")
        attach_logs = self.config.get_bool("ATTACH_LOG_ON_ERROR", False)
        send_attachment = self.config.get_bool("MAIL_SEND_ATTACHMENT", True)

        include_screenshots = self.config.get_bool("MAIL_INCLUDE_SCREENSHOTS", True)
        include_log_snippet = self.config.get_bool("MAIL_INCLUDE_LOG_SNIPPET", True)

        if not to_addrs:
            self.logger.warning("MAIL_TO vacío → no se enviará correo.")
            return

        # Enriquecer el cuerpo si hay error y está habilitado el snippet
        if include_log_snippet and ("ERROR" in subject.upper() or "FALLÓ" in body.upper()):
            body += "\n\n--- ÚLTIMAS LÍNEAS DEL REGISTRO ---\n"
            body += self._get_log_snippet()

        # Preparar adjuntos
        final_attachments = attachments or []
        
        # Adjuntar capturas de pantalla si existen y está habilitado
        if include_screenshots:
            screenshot_dir = "logs/screenshots"
            if os.path.exists(screenshot_dir):
                try:
                    screenshots = [os.path.join(screenshot_dir, f) for f in os.listdir(screenshot_dir) if f.endswith(".png")]
                    if screenshots:
                        # Solo las 2 más recientes para no saturar
                        screenshots.sort(key=os.path.getmtime, reverse=True)
                        final_attachments.extend(screenshots[:2])
                except:
                    pass

        if attach_logs and ("ERROR" in subject.upper() or "FALLÓ" in body.upper()):
            log_dir = self.config.get("LOG_DIR", "logs")
            log_file = self.config.get("LOG_FILE", "botexcel.log")
            log_path = os.path.abspath(os.path.join(log_dir, log_file))
            if os.path.exists(log_path):
                final_attachments.append(log_path)
            else:
                self.logger.warning(f"No existe el log para adjuntar: {log_path}")

        if use_outlook:
            # Enviar usando Outlook Desktop
            try:
                pythoncom.CoInitialize()
                outlook = win32.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.Subject = subject
                mail.Body = body
                mail.To = ";".join(to_addrs)
                if from_addr:
                    mail.SentOnBehalfOfName = from_addr

                # Usar conjunto para evitar duplicados
                for path in set(final_attachments):
                    try:
                        mail.Attachments.Add(path)
                        self.logger.info(f"Adjuntando archivo: {path}")
                    except Exception as e:
                        self.logger.warning(f"No se pudo adjuntar {path}: {e}")

                mail.Send()
                self.logger.info("Correo enviado correctamente vía Outlook Desktop.")
            except Exception as e:
                self.logger.warning(f"Fallo al enviar correo vía Outlook Desktop: {e}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        else:
            # Enviar usando SMTP
            smtp_server = self.config.get("MAIL_SMTP_SERVER")
            smtp_port = int(self.config.get("MAIL_SMTP_PORT", 587))
            username = self.config.get("MAIL_FROM")
            password = self.config.get("MAIL_PASSWORD")
            use_tls = self.config.get_bool("MAIL_USE_TLS", True)

            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = from_addr
            msg["To"] = ", ".join(to_addrs)
            msg.set_content(body)

            for path in set(final_attachments):
                try:
                    with open(path, "rb") as f:
                        data = f.read()
                        filename = os.path.basename(path)
                        msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=filename)
                    self.logger.info(f"Adjuntando archivo: {path}")
                except Exception as e:
                    self.logger.warning(f"No se pudo adjuntar {path}: {e}")

            try:
                server = smtplib.SMTP(smtp_server, smtp_port)
                if use_tls:
                    server.starttls()
                server.login(username, password)
                server.send_message(msg)
                server.quit()
                self.logger.info("Correo enviado correctamente vía SMTP.")
            except Exception as e:
                self.logger.warning(f"Error enviando correo vía SMTP: {e}")
