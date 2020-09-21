

__all__ = (
    "BaseExcelFormulaError",
    "ExcelFormulaError",
    "ExcelNameError",
    "ExcelValueError"
    )

class BaseExcelFormulaError(Exception):
    pass


class ExcelFormulaError(BaseExcelFormulaError):
    pass


class ExcelNameError(NameError, BaseExcelFormulaError):
    pass


class ExcelValueError(ValueError, BaseExcelFormulaError):
    pass
