# domain/refresh_job.py

from domain.exceptions import RefreshJobError, ExcelGatewayError


class RefreshJob:
    """
    Representa un trabajo de refresh de un archivo Excel.
    Encapsula la lógica de ejecución contra el gateway de Excel.
    """

    def __init__(self, excel_path: str):
        self.excel_path = excel_path

    def execute(self, excel_gateway):
        """
        Ejecuta el refresh usando el gateway.
        Lanza RefreshJobError si algo falla.
        """
        try:
            return excel_gateway.refresh_file(self.excel_path)
        except ExcelGatewayError as e:
            raise RefreshJobError(str(e))
