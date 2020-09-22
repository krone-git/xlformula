# ./lib/builtin/_function.py

"""
Module defines and dynamically generates builtin 'ExcelFunction' classes
to emulate those found in Excel formulas.

---------------------------------------------------------------------

Users should not import directly from this module.
All function classes defined here are imported into the 'builtin' package
and 'lib' package and then into the top-level 'xlformula' package.

Function classes can be imported from either package with the
'from [package] import *' statement.
"""

from collections.abc import Iterable
import math, statistics, random
try:
    from math import prod       # Product function for Python 3.8 and above;
except ImportError as e:
    from functools import reduce
    from operator import mul

    def prod(iterable, *, start: int=1) -> int:
        # Product function for Python 3.7 and below;
        return reduce(mul, iterable, start)

# Import base function class, default 'get_value' method to raise
# 'NotImplementedError' when called, and 'get_value' constant for use as
# class namespace argument;
from ..function import ExcelFunction, _notimplemented, GET_VALUE as GET
# Imports '__requiredarguments__', '__optionalarguments__', and
# '__inheritsarguments__' constants for use as class namespace arguments;
from ..argument import REQUIRED_ARGUMENTS as REQ, \
                        OPTIONAL_ARGUMENTS as OPT, \
                        INHERIT_ARGUMENTS as INH


BASE = "bases"          # Define 'bases' constant for storing ;
DOC = "__doc__"         # Define '__doc__' constant for storing '__doc__'
                        # string function class namespace argument;

# Commonly used sets of required and option arguments;
_NUMBER = {
    REQ: ("number",)
    }
_NUMBER_RANGE = {
    REQ: ("number1",),
    OPT: ("number2", ...)
    }

_LOGIC = {
    REQ: ("logical",)
    }
_LOGIC_RANGE = {
    REQ: ("logical1",),
    OPT: ("logical2", ...,),
    }

_TEXT = {
    REQ: ("text",)
    }
_TEXT_RANGE = {
    REQ: ("text1",),
    OPT: ("text2", ...,),
    }

_VALUE_RANGE = {
    REQ: ("value1",),
    OPT: ("value2", ...)
    }


def _getvalues(self):
    """Common use method for collecting function call arguments."""
    return (
        arg.get_value() for arg in self._arguments
        )

# Define dictionary to store function class variables for use in
# dynamically generating function classes;
_BUILTIN_FUNCTIONS = {
    # "FUNCTION": {             # Key defines function's class name;
    #     REQ: ("",),           # Defines function's required arguments.
                                # Can be single 'str' or tuple of 'str';
    #     OPT: ("",),           # Defines function's optional arguments.
                                # Can be single 'str' or tuple of 'str';
    #     INH: False,           # If 'True', function will inherit
                                # required and optional arguments from bases.
                                # Otherwise, function will override arguments;
    #     BASE: ("",),          # Parent function class bases.
                                # Can be single 'str' or tuple of 'str';
    #     GET: _notimplemented, # Method for function's 'get_value' method;
    #     DOC: ""               # Function class '__doc__' string;
    #     },


    # Logical functions
    "AND": {
        **_LOGIC_RANGE,
        GET: (
            lambda self: all(self.get_values())
            ),
        },
    "IF": {
        REQ: ("logical_test", "value_if_true"),
        OPT: ("value_if_false",),
        GET: (
                lambda self: self._arguments[1].get_value() \
                    if self._arguments[0].get_value() \
                    else self._arguments[2].get_value()
            )
        },
    "NOT": {
        **_LOGIC,
        GET: (lambda self: not self.get_value()),
        },
    "OR": {
        **_LOGIC_RANGE,
        GET: (
            lambda self: any(self.get_values())
            ),
        },
    "TRUE":{
        GET: (
            lambda self: True
            )
        },


    # Arithmetic functions
    "ABS": {
        **_NUMBER,
        GET: (
            lambda self: abs(self.get_value())
            )
        },
    "CEILING": {
        REQ: ("number", "significance"),
        OPT: (),
        GET: (
            lambda self: math.ceil(self.get_value())
            ),
        },
    # "PI": {
    #
    #     },
    # "POWER": {
    #
    #     },
    "PRODUCT": {
        **_NUMBER_RANGE,
        GET: (
            lambda self: prod(self.get_values())
            )
        },
    "RAND": {
        GET: (
            lambda self: random.random()
            )
        },
    "SUM": {
        **_NUMBER_RANGE,
        GET: (
            lambda self: sum(self.get_values())
            )
        },


    # Statistical functions
    "AVERAGE": {
        **_NUMBER_RANGE,
        GET: (
            lambda self: statistics.mean(self.get_values())
            )
        },
    "COUNT": {
        **_VALUE_RANGE
        },
    "COUNTA": {
        **_VALUE_RANGE
        },
    "COUNTBLANK": {
        REQ: ("range",)
        },
    "COUNTIF": {
        REQ: ("range", "criteria")
        },
    # "SMALL": {
    #
    #     },


    # String functions
    "CHAR": {
        REQ: ( "number",),
        },
    "CONCATENATE": {
        **_TEXT_RANGE,
        GET: (
            lambda self: "".join(str(i) for i in self.get_values())
            ),
        },
    "LEN" :{
        **_TEXT,
        GET: (
            lambda self: len(self.get_value())
            )
        },
    # "SUBSTITUTE": {
    #
    #     },
    # "TRIM": {
    #
    #     },
    # "VALUE": {
    #
    #     },


    # Date and Time functions
    "DATE": {
        REQ: ("year", "month", "day")
        },
    "DATEVALUE": {
        REQ: ("date_text",)
        },
    "DAY": {
        REQ: ("serial_number",)
        },
    # "NOW": {
    #
    #     },
    "TIME": {
        REQ: ("hour", "minute", "second")
        },
    "TIMEVALUE": {

        },
    # "TODAY": {
    #
    #     },


    # Information functions
    "CELL": {
        REQ: ("info_type",),
        OPT: ("reference"),
        },
    # "NA": {
    #
    #     }


    # Reference functions
    "ADDRESS": {
        REQ: ("row_num", "column_num"),
        OPT: ("abs_num",)
        },
    "COLUMN": {
        OPT: ("reference",)
        },
    "COLUMNS": {
        REQ: ("array",)
        },
    # "OFFSET": {
    #
    #     },
    "ROW": {
        OPT: ("reference",)
        },

    }

__all__ = (                         # Define '__all__' for implicit '*' imports;
    *_BUILTIN_FUNCTIONS.keys(),     # Only generated function classes should
                                    # be imported from namespace;
    "FUNCTIONS"
    )


FUNCTIONS = dict()                      # Defines dictionary to store
                                        # generates functions classes for
                                        # use in inheritance;
_vars = vars()                          # Stores namespace 'vars()' to avoid
                                        # multiple calls to 'vars()';
for k, v in _BUILTIN_FUNCTIONS.items():
    name = k.upper()                    # Cast function class name to uppercase;
    bases = v.pop(BASE, ())             # Pulls base class names;
    v.setdefault(GET, _notimplemented)  # Sets function 'get_value' method
                                        # to raise 'NotImplementedError' if
                                        # method is not defined;

    if isinstance(bases, str) or not isinstance(bases, Iterable):
        # If 'bases' is not an iterable, casts 'bases' to tuple;
        # "arg" -> ("arg",)
        bases = (str(bases),)

    # Pulls previously generated function classes to as base for inheritance;
    # Throws 'KeyError' if base function name not found to prevent incorrect
    # inheritance;
    bases = (
        FUNCTIONS[str(base).upper()] for base in bases if base
        )

    # Dynamically generates function class;
    func = type(name, (ExcelFunction, *bases), v)
    # Adds generated function class to namespace 'vars()' and 'FUNCTIONS'
    # dictionary for use as an inheritable base class;
    _vars.sedefault(name, func)
    FUNCTIONS[name] = func

FUNCTIONS = frozenset(FUNCTIONS.keys())

del _BUILTIN_FUNCTIONS, _vars, k, v     # Delete variables to prevent explicit
                                        # imports;
