from abc import ABCMeta, abstractmethod

from .argument import ExcelArgumentHandler, ExcelArgumentHandlerType
from .composite import _format_argument
from .call import ExcelFormulaCall


__all__ = (
    "ExcelFormula",
    )


class ExcelFormulaArgumentHandlerType(ExcelArgumentHandlerType):
    def __repr__(cls):
        return f"<{ExcelFormula.__name__}> {super().__repr__()}"


class ExcelFormulaType(ExcelFormulaArgumentHandlerType,
                        ABCMeta):
    pass


class ExcelFormula(ExcelArgumentHandler,
                    metaclass=ExcelFormulaType):
    @classmethod
    def _handle_arguments(cls, args):
        args = (
            _format_argument(arg) for arg in args
            )
        args = cls.formulate(cls, *args)
        return ExcelFormulaCall(cls, args)

    @abstractmethod
    def formulate(self, *args):
        raise NotImplementedError

    def get_value(self):
        return (arg.get_value() for arg in self._arguments)
