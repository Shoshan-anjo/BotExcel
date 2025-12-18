# domain/exceptions.py

class RefreshJobError(Exception):
    pass

class ConfigError(Exception):
    pass

class ExcelGatewayError(Exception):
    pass

class EmailNotificationError(Exception):
    pass

class EmailConfigError(Exception):
    pass
