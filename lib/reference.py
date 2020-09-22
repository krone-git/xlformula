# ./lib/reference.py

"""
Module defines 'ExcelReference' classes for working with Excel range
references.

---------------------------------------------------------------------

Users should not import directly from this module.
All relevant classes and constants defined here are imported into the
'lib' package and then into the top-level 'xlformula' package.

Classes and constants can be imported from either package with the
'from [package] import *' statement.

All abstract and meta classes defined here are imported into the 'abc'
module.
"""

from string import ascii_uppercase
import re

# Import base 'ExcelStringBuilder' class for 'ExcelReference' class
# inheritance;
from .composite import ExcelStringBuilder


_REFERENCE_VALUES = (       # Defines tuple of reference_type constant names
                            # to be dynamically generated in the namespace;
    "ABS_REF",              # Fully absolute reference: 1 (e.g. $A$1);
    "ABS_ROW",              # Row-only absolute reference: 2 (e.g. A$1);
    "ABS_COL",              # Column-only absolute reference: 3 (e.g. $A1);
    "REL_REF",              # Fully relative reference: 4 (e.g. A1);
    )

_vars = vars()                              # Stores namespace 'vars()' to
                                            # avoid multiple calls to 'vars()';

for i, v in enumerate(_REFERENCE_VALUES):
    # Dynamically generates constant with uppercase varname and sets to 'int';
    _vars[v.upper()] = i + 1                # Set from index 0 to index 1;

__all__ = (                                     # Defines '__all__' for
                                                # implicit '*' imports;
    *(i.upper() for i in _REFERENCE_VALUES),    # Allows for import of
                                                # reference type constants;
    "ExcelReference",                           # Allows for import of
                                                # 'ExcelReference' class;
    "ExcelRangeReference"                       # Allows for import of
                                                # 'ExcelCellReference' class;
    )

del _REFERENCE_VALUES, _vars, i, v      # Delete variables to prevent
                                        # explicit imports;

# Defines regular expression patterns for parsing Excel refernce strings;
_WORKBOOK_REFERENCE_PATTERN = re.compile(
    "\[.*\..*\]"                            # Defines compiled 're' pattern
                                            # to return the workbook filename
                                            # with '[]'s contained within a
                                            # reference;
    )
_SHEET_REFERENCE_PATTERN = re.compile(
    "\'.+?\'!"                              # Defines compiled 're' pattern
                                            # to return the sheet/tab name
                                            # with "''!" contained within a
                                            # reference;
    )
_RANGE_REFERENCE_PATTERN = re.compile(
    "\'![\$:a-zA-Z0-9]+"                    # Defines compiled 're' pattern
                                            # to return the range with ':'
                                            # contained within a reference;
    )
_CELL_REFERENCE_PATTERN = re.compile(
    "\$?[a-zA-Z]+|\$?[0-9]+"                # Defines compiled 're' pattern
                                            # to return the cells within a
                                            # range contained within a
                                            # reference;
    )
# String cleaning patterns
_WORKBOOK_NAME_PATTERN = re.compile(
    "[^\[\]]+[^\[\]]"                       # Defines compiled 're' pattern
                                            # to return the name only from
                                            # a workbook found within a
                                            # reference;
    )
_SHEET_NAME_PATTERN = re.compile(
    "[^\']+[^\']"                           # Defines compiled 're' pattern
                                            # to return the name only from
                                            # a sheet contained within a
                                            # reference;
    )


def _alpha_recursion(value: int, string: str) -> str:
    """
    Recursive function for use in the '_alpha_column' function.
    Calculates base 26, alphabetic value of 'value' and appends it to
    'string', then calls itself with 'string' and floor division of
    base 26 and 'value'.
    """
    base = len(ascii_uppercase)             # Define base for division (26);
    mod, div = value % base, value // base  # Calculate module and floor
                                            # divison of 'value' and 'base';
    string = ascii_uppercase[mod] + string  # Append alphabetic character
                                            # found at index 'mod' to string
    if div < 1:
        # If there is no remaining floor end recursion and return string.
        # Otherwise, continue recursion with floor division;
        return string
    else:
        div -= 1                            # Set from index 1 to index 0;
        return _alpha_recursion(div, string)

def _alpha_column(value: int) -> str:
    """
    Initiates recursive '_alpha_recursion' function to build string of
    alphabetic characters with base 26 with base 10 value 'col'.
    """
    if isinstance(value, str) and value.upper() in ascii_uppercase:
        # If 'column' is already a 'str' return 'column' cast to uppercase;
        return value.upper()
    elif not isinstance(value, int):
        # If 'column' is not an 'int' or 'str' raise 'TypeError';
        raise TypeError("Column index type must be 'int'.")

    value -= 1                          # Set from index 1 to index 0;
    return _alpha_recursion(value, "")  # Begin recursion and return result;

def _numeric_column(column: str) -> int:
    """
    Calculates base 10, numeric value from base 26, alphabetic string 'column'.
    """
    if isinstance(column, int):
        # If 'column' is already an 'int' return 'column';
        return column

    column = column.upper()[::-1]   # Cast 'column' to uppercase and reverse it;
    base = len(ascii_uppercase)     # Define base (26);
    num = 0                         # Define initial base 10 value;

    for index, char in enumerate(column):
        # Increment 'num' by the product of the index 1 position
        # of the character in the alphabet and base 26 to the power of the
        # character's position in 'column';
        num += (base ** index) * (ascii_uppercase.index(char) + 1)
    return num                      # Return base 10 result;


def _remove_workbook_reference(string: str) -> str:
    """Returns reference with workbook section omitted."""
    return string.replace(f"{_parse_workbook_reference(string)}", "")

def _remove_sheet_reference(string: str) -> str:
    """Returns reference with sheet section omitted."""
    return string.replace(f"{_parse_sheet_reference(string)}", "")

def _remove_range_reference(string: str) -> str:
    """Returns reference with range section omitted."""
    return string.replace(f"{_parse_range_reference(string)}", "")

def _parse_workbook_reference(string: str) -> str:
    """Finds and returns workbook section of the reference."""
    return _WORKBOOK_REFERENCE_PATTERN.search(string).group()

def _parse_sheet_reference(string: str) -> str:
    """Finds and returns sheet section of the reference."""
    return _SHEET_REFERENCE_PATTERN.search(string).group()

def _parse_range_reference(string: str) -> str:
    """Finds, and returns cleaned range section of the reference."""

    # NOTE: 9/20/20 Brandon Krone
    # Needs fixing; pattern returns '! from end of sheet name
    # Hot fixed to shave first two characters from head of string ([2:]);
    # Needs fix in regex pattern to eliminate extra characters cleanly;
    return _RANGE_REFERENCE_PATTERN.search(string).group()[2:]

def _parse_workbook_name(string: str) -> str:
    """Cleans and returns found workbook section of reference."""
    reference = _parse_workbook_reference(string)
    return _WORKBOOK_NAME_PATTERN.search(reference).group()

def _parse_sheet_name(string: str) -> str:
    """Cleans and returns found sheet section of reference."""
    reference = _parse_sheet_reference(string)
    return _SHEET_NAME_PATTERN.search(reference).group()

def _parse_range_string(string: str) -> str:
    """Returns tuple of cell references contained within a range reference."""
    return tuple(_parse_range_reference(string).split(":"))

def _parse_cell_string(string: str) -> str:
    """Returns tuple of column and row contained within a cell reference."""
    return tuple(_CELL_REFERENCE_PATTERN.findall(string))

def _parse_reference_arguments(string: str) -> str:
    ###
    raise NotImplementedError

def _parse_range_arguments(string: str) -> str:
    ###
    raise NotImplementedError

def _parse_cell_arguments(string: str) -> str:
    ###
    raise NotImplementedError


def _build_single_cell_string(row: int, col: int, ref: int=REL_REF) -> str:
    """
    Constructs and returns cell reference string from row and column indexes
    'col' and 'row' with reference type 'ref'.
    """
    row_string = str(row) if row else ""            # Set to "" if 'row' is 0;
    if row_string and ref in (ABS_REF, ABS_ROW):
        # If reference type is 1 or 2, append absolute reference character '$';
        row_string = f"${row_string}"

    col_string = _alpha_column(col)                 # Find base 26 alphabetic
                                                    # value if base 10 column
                                                    # index 'col';
    if col_string and ref in (ABS_REF, ABS_COL):
        # If reference type is 1 or 3, append absolute reference character '$';
        col_string = f"${row_string}"

    return "".join((col_string, row_string))        # Join and return column
                                                    # and row strings;

def _build_reference_string(row: int, col: int,
                            ref: int=ABS_REF,
                            *, sheet: str=None,
                            workbook: str=None,
                            r_row: int=None,
                            r_col: int=None,
                            r_ref: int=REL_REF) -> str:
    """
    Constructs and returns full reference string with workbook name
    'workbook', sheet name 'sheet', and cell refernce from column index 'col',
    row index 'row' and reference type 'ref'. Constructs with a range
    reference if range indexes 'r_row' and 'r_col' and reference type 'r_ref'
    parameters are provided.
    """
    wb_string = f"[{workbook}]" if workbook else "" # Append workbook reference
                                                    # identifiers to workbook
                                                    # name, if it exists.
                                                    # Otherwise set to empty
                                                    # string;
    sheet_string = f"'{sheet}'!" if sheet else ""   # Append sheet reference
                                                    # identifiers to sheet
                                                    # name, if it exists.
                                                    # Otherwise set to empty
                                                    # string;

    # Construct reference string for first cell in range
    ref_string = _build_single_cell_string(row, col, ref)
    if (r_row or r_col) and r_ref:
        # If value for range values 'r_row' or 'r_col' exist,
        # Construct reference string for second cell in range.
        range_ref_string = _build_single_cell_string(r_row, r_col, r_ref)
    else:
        range_ref_string = ""

    if range_ref_string:
        # If second cell reference is not an empty string,
        # join both reference strings into range reference string;
        ref_string = ":".join((ref_string, range_ref_string))

    # Join and return workbook, sheet, and range reference strings;
    return "".join((wb_string, sheet_string, ref_string))


class ExcelReferenceType(ExcelStringBuilder):
    # Intended to be imported for type checking of 'ExcelReference' classes;
    # Imports into xlformula.abc package;
    """Abstract base class for 'ExcelReference' classes."""
    pass


class ExcelReference(ExcelReferenceType):
    """
    Simple implementaion for working with excel references.
    Compiles with 'name' only when compiling formulas.
    Useful when including Excel named ranges within a formula.

    ------

    Passing an 'ExcelReference' object to an 'ExcelFunction' or 'ExcelFormula'
    instead of a 'str' will bypass '""' being added to its string when
    compiling formulas.
        e.g.
            str 'Named_Range' -> '"Named_Range"'
            ExcelReference 'Named_Range' -> 'Named_Range'
    """
    def __init__(self, name: str):
        self._name = name

    @property
    def name(self) -> str:
        "Returns reference's name."
        return self._name

    def compile(self) -> str:
        "Returns reference's formula compiling string."
        return self._name


class ExcelRangeReference(ExcelReferenceType):
    """
    Advanced implementaion for working with excel range and cell references.
    Constructs excel range reference from row and column indexes with
    with support for workbook and sheet(tab) names.
        e.g.
            wb_ref = ExcelRangeReference(
                1, 1, ref=1, sheet="Sheet1", workbook="Book1.xlsx",
                r_row=2, r_col=1, r_ref=4
                )
            wb_ref.compile() -> "[Book1.xlsx]'Sheet1'!$A$1:A2"

            range_ref = ExcelRangeReference(
                1, 1, ref=2, r_row=2, r_col=1, r_ref=3
                )
            range_ref.compile() -> "A$1:$A2"

    """
    def __init__(self, row: int=1, col: int=1, ref: int=ABS_REF, *,
                    sheet: str=None, tab: str=None, workbook: str=None,
                    r_row: int=None, r_col: int=None, r_ref: int=None):
        self._workbook = workbook
        # Keywords 'sheet' and 'tab' represent the same value;
        self._sheet = tab if sheet is None and tab is not None else sheet
        # Stores range reference parameters as tuples;
        self._reference = (row, col, ref)
        self._range = (r_row, r_col, r_ref)
    # Append note to '__init__' '__doc__' string;
    __init__.__doc__ = "\n\n".join((
        object.__init__.__doc__,
        "* Keywords 'sheet' and 'tab' can be used interchangeably."
        ))

    @property
    def workbook(self) -> str:
        """Returns references's workbook name."""
        return self._workbook

    @property
    def sheetname(self) -> str:
        """Returns reference's sheet name."""
        return self._sheet

    tab = sheet = sheetname         # Set 'tab' and 'sheet' property to be
                                    # interchangeable with 'sheetname';
    @property
    def range_reference(self) -> str:
        """
        Returns reference's range as a string without its workbook or
        sheet name.
        """
        # Unpack range reference parameters;
        row, col, ref = self._reference
        r_row, r_col, r_ref = self._range
        # Construct and return full reference string;
        return _build_reference_string(
            row,
            col,
            ref=ref,
            r_row=r_row,
            r_col=r_col,
            r_ref=r_ref
            )

    @property
    def reference(self) -> str:
        """Returns reference as a string."""
        # Class specific 'compile' method as a property;
        return self.compile()

    def compile(self) -> str:
        # Unpack range reference parameters;
        row, col, ref = self._reference
        r_row, r_col, r_ref = self._range
        # Construct and return full reference string;
        return _build_reference_string(
            row,
            col,
            sheet=self._sheet,
            workbook=self._workbook,
            ref=ref,
            r_row=r_row,
            r_col = r_col,
            r_ref=r_ref
            )
