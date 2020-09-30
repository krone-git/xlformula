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
from ..reference import *


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


class IfFirstColumnRow(ExcelFormula):
    """
    """
    __requiredarguments__ = (
        "reference",
        "value_if_true",
        "value_if_false"
        )
    __optionalarguments__ = ("reference_type",)

    def formulate(self, reference, if_true, if_false, ref=ABS_REF):
        column, row = get_column_row(reference.get_value())
        relative = ExcelRangeReference(row, column, ref=REL_REF)
        absolute = ExcelRangeReference(row, column, ref=ref)
        return (
            IF(relative==absolute, if_true, if_false),
            )


class IfModulo(ExcelFormula):
    """
    """
    __requiredarguments__ = (
        "logical",
        "modulo",
        "value_if_true",
        "value_if_false"
        )
    __optionalarguments__ = ("if_modulo_equals",)
    
    def formulate(self, logical, module, if_true, if_false, remainder=0):
        return (
            IF(MOD(logical, module) == remainder, if_true, if_false),
            )

    
class IfModuloChain(ExcelFormula):
    """
    """
    __requiredarguments__ = (
        "logical_test",
        "value_if_true1",
        "value_if_false_final"
        )
    __optionalarguments__ = (
        "value_if_true2",
        ...,
        "start"
        )
    
    def formulate(self, logical, if_true1, if_false_final, *args, start=0):
        args = (
            if_true1,
            if_false_final,
            *args
            )
        if_false_final = args[-1]
        args = args[:-1]
        mod = len(args)

        for i in reversed(range(mod)):
            if i > len(args) - 2:
                _if_false = if_false_final
            else:
                _if_false = modulo 
            modulo = IfModulo(logical, mod, args[i], _if_false, i)

        return (modulo,)

class BuildFilterString(ExcelFormula):
    """
    """
    __requiredarguments__ = ("starting_reference", "delimiter")
    __optionalarguments__ = ("double_delimiter",)

    def formulate(self, ref, delimiter, double_delimiter=None):
        if double_delimiter is None:
            if isinstance(delimiter, ExcelArgument):
                _delimiter = delimiter.get_value()
            else:
                _delimiter = delimiter
            double_delimiter = ExcelArgument(_delimiter * 2)
            
        
        return (
            ,
            )


__all__ = (
    *(
        k for k, v in vars().items() \
        if isinstance(v, type) and issubclass(v, ExcelFormula)
        ),
    )
