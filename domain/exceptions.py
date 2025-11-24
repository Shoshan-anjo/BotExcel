# domain/exceptions.py


class ConfigError(Exception):
    """Errores relacionados con configuración (.env, rutas, etc.)."""
    pass


class ExcelGatewayError(Exception):
    """Errores producidos al interactuar con Excel via COM."""
    pass


class EmailNotificationError(Exception):
    """Errores al enviar notificaciones por correo."""
    pass


class RefreshJobError(Exception):
    """Errores de alto nivel al ejecutar un trabajo de refresh."""
    pass
