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
from ..composite import ExcelArgument
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
    Formula returns 'value_if_true' if the current cell in the algorithm is
    the first cell in the algorithm, otherwise 'value_if_false' is returned.

    Relative form of reference 'ref' is compared against absolute form.
    """
    __requiredarguments__ = (
        "reference",
        "value_if_true",
        "value_if_false"
        )
    __optionalarguments__ = ("reference_type",)

    def formulate(self, reference, if_true, if_false, ref=ABS_REF):
        column, row = get_column_row(reference.reference)
        relative = ExcelRangeReference(row, column, ref=REL_REF)
        absolute = ExcelRangeReference(row, column, ref=ref)

        if ref in (ABS_ROW, ABS_REF, REL_REF):
            func = ROW
        elif ref in (ABS_COL):
            func = COLUMN
        else:
            raise ValueError    ######

        return (
            IF(
                func(relative) == func(absolute),
                if_true,
                if_false
                ),
            )


class IfModulo(ExcelFormula):
    """
    Base component for IfModuleChain.

    Formula returns 'value_if_true' if the modulo of 'logical' and 'modulo'
    equals 'remainder' (0 by default), otherwise 'value_if_false' is returned.
    """
    __requiredarguments__ = (
        "logical",
        "modulo",
        "value_if_true",
        "value_if_false"
        )
    __optionalarguments__ = ("if_modulo_equals",)

    def formulate(self, logical, modulo, if_true, if_false, remainder=0):
        return (
            IF(MOD(logical, modulo) == remainder, if_true, if_false),
            )


class IfModuloChain(ExcelFormula):
    """
    Formula performs a modulo check on 'logical' and returns the 'if_true'
    value for the first check that returns 'True'. If none of the modulo
    checks return 'True', 'if_false_final' is returned.

    Formula is intended to facilitate the creation of dynamic, two-dimensional
    table sections, which can be repeated indefinitely.

    --------------------------------------------------------------------------

    Any number of 'if_true' values can be passed. The final value will be
    assumed as 'if_false_final'.
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
    Formula creates a volatile filtering algorithm that will filter a list of
    values in real-time without the need for the user to refresh the filter
    as with Tables, Pivot Tables and basic and Advanced Filters.

    Fromula also does not hide rows as with basic and Advanced Filters, so
    additional rows, tables, and formulas can be included in the same sheet.

    Use in combination with the 'SplitFilterString' Formula to retrieve
    filtered values. 'FilterString' algorithm should be contained in initial
    table to be sorted, then a new table should be created to split the
    resulting 'FilterString' and use INDEX(MATCH(...)) to transfer values from
    initial table to the filtered table.
    """
    __requiredarguments__ = (
        "logical",
        "cell_reference",
        "start_reference",
        "delimiter"
        )
    __optionalarguments__ = (
        "reference_type",
        "reverse"
        )

    def formulate(self, logical, reference, start_reference, delimiter,
                    ref=ABS_REF, reverse=False):
        null_string = ExcelArgument("")

        print(start_reference, type(start_reference))
        column, row = get_column_row(start_reference.reference)
        if ref in (ABS_ROW, ABS_REF, REL_REF):
            row -= 1
        elif ref in (ABS_COL):
            column -= 1
        else:
            raise ValueError    ######

        previous = ExcelRangeReference(row, column, ref=REL_REF)

        head = IfFirstColumnRow(
            start_reference,
            null_string,
            previous,
            ref=ref
            )
        tail = IF(
            logical,
            CONCATENATE(delmiter, reference, delimiter),
            null_string
            )

        return (
            CONCATENATE(
                tail if reverse else head,
                head if reverse else tail
                ),
            )

class BuildUniqueFilterString(ExcelFormula):
    """
    Formula creates a volatile filtering algorithm that will filter a list of
    values in real-time without the need for the user to refresh the filter
    as with Tables, Pivot Tables and basic and Advanced Filters.

    Fromula also does not hide rows as with basic and Advanced Filters, so
    additional rows, tables, and formulas can be included in the same sheet.

    Use in combination with the 'SplitFilterString' Formula to retrieve
    filtered values. 'FilterString' algorithm should be contained in initial
    initial table to the filtered table.
    table to be sorted, then a new table should be created to split the
    resulting 'FilterString' and use INDEX(MATCH(..)) to transfer values from
    """
    __requiredarguments__ = ("starting_reference", "delimiter")
    __optionalarguments__ = ("double_delimiter")

    def formulate(self, ref, delimiter, double_delimiter=None):
        null_string = ExcelArgument("")

        column, row = get_column_row(start_reference.reference)
        if ref in (ABS_ROW, ABS_REF, REL_REF):
            row -= 1
        elif ref in (ABS_COL):
            column -= 1
        else:
            raise ValueError    ######

        previous = ExcelRangeReference(row, column, ref=REL_REF)
        delimited_value = CONCATENATE(delmiter, reference, delimiter)

        head = IfFirstColumnRow(
            start_reference,
            null_string,
            previous,
            ref=ref
            )
        tail = IF(
            AND(
                logical,
                ISERROR(
                    FIND(
                        delimited_value,
                        previous
                        )
                    )
                ),
            delimited_value,
            null_string
            )

        return (
            CONCATENATE(
                tail if reverse else head,
                head if reverse else tail
                ),
            )

class SplitFilterString(ExcelFormula):
    """
    """

    def formulate():
        if double_delimiter is None:
            if isinstance(delimiter, ExcelArgument):
                _delimiter = delimiter.get_value()
            else:
                _delimiter = delimiter
            double_delimiter = ExcelArgument(_delimiter * 2)

        pass


__all__ = (
    *(
        k for k, v in vars().items() \
        if isinstance(v, type) and issubclass(v, ExcelFormula)
        ),
    )
