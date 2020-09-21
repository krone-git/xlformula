# ./lib/builtin/_formula.py

"""
Module defines builtin 'ExcelFormula' classes for creating commonly
used combinations of functions.

---------------------------------------------------------------------

Users should not import directly from this module.
All formula classes defined here are imported into the 'builtin' package
and 'lib' package and then into the top-level 'xlformula' package.

Formula classes can be imported from either package with the
'from [package] import *' statement.
"""

# Import all builtin function classes, 'ExcelFormula' base class,
# and 'ExcelArgument' class;
from ._function import *
from ..formula import ExcelFormula
from ..composite import ExcelArgument as Arg


class BlankIfBlank(ExcelFormula):
    """
    Formula returns "" if reference value is "", else it returns
    value defined in 'else'.
        e.g.
            A1="" -> ""
            or
            A1=1 -> [else]
    """
    __requiredarguments__ = ("reference", "else")

    def formulate(self, reference: str, _else) -> tuple:
        return (
            IF(reference=="", "", _else),
            )


class GetColumnLetter(ExcelFormula):
    """
    Formula returns the column letter notation of single cell reference.
        e.g.
            $AB$1 -> AB
            or
            'Sheet1'!$ABC$123 -> ABC
    """
    __requiredarguments__ = ("reference",)

    def formulate(self, reference: str) -> tuple:
        return (
            SUBSTITUTE(ADDRESS(ROW(), COLUMN(reference), 4), ROW(), "")
            )

__all__ = (
    *(
        k for k, v in vars().items() \
        if isinstance(v, type) and issubclass(v, ExcelFormula)
        ),
    )
