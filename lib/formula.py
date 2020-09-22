# ./lib/formula.py

"""
Module defines the abstract 'ExcelFormula' class for defining and working
with user-defined Excel formulas.

---------------------------------------------------------------------

Users should not import directly from this module.
All relevant classes and constants defined here are imported into the
'lib' package and then into the top-level 'xlformula' package.

Classes and constants can be imported from either package with the
'from [package] import *' statement.

All abstract and meta classes defined here are imported into the 'abc'
module.
"""

from abc import ABCMeta, abstractmethod

from .argument import ExcelArgumentHandler, ExcelArgumentHandlerType
from .composite import _format_argument
from .call import ExcelFormulaCall


__all__ = (             # Defines '__all__' for implicit '*' imports;
    "ExcelFormula",     # Only the 'ExcelFormula' class should be imported
                        # form this module;
    )


class ExcelFormulaArgumentHandlerType(ExcelArgumentHandlerType):
    """
    Abstract class for handling and enforcing required and optional
    arguments for 'ExcelFormula' subclasses.
    """
    def __repr__(cls):
        # Return the classname with the representation atring of the
        # parent class;
        return f"<{ExcelFormula.__name__}> {super().__repr__()}"


class ExcelFormulaType(ExcelFormulaArgumentHandlerType,
                        ABCMeta):
    """Abstract class for the 'ExcelFormula' class."""
    pass


class ExcelFormula(ExcelArgumentHandler,
                    metaclass=ExcelFormulaType):
    """
    Base implementation for working with Excel formulas.


    """
    @classmethod
    def _handle_arguments(cls, args: tuple):
        """"""
        args = (
            _format_argument(arg) for arg in args
            )
        args = cls.formulate(cls, *args)
        return ExcelFormulaCall(cls, args)

    @abstractmethod
    def formulate(self, *args):
        #
        raise NotImplementedError

    def get_value(self):
        """"""
        return (arg.get_value() for arg in self._arguments)
