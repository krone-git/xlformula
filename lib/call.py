# ./lib/call.py

"""
Module defines 'ExcelCall' classes for storing and handling individual
sets of arguments passed to a call of the 'ExcelFunction' and 'ExcelFormula'
classes.

---------------------------------------------------------------------

Users should not import directly from this module.
All relevant classes and constants defined here are imported into the
'lib' package and then into the top-level 'xlformula' package.

Classes and constants can be imported from either package with the
'from [package] import *' statement.

All abstract and meta classes defined here are imported into the 'abc'
package.
"""

from abc import ABCMeta, abstractmethod
from collections.abc import Sequence
from functools import wraps

# Import 'ExcelCompositeType' and 'ExcelComposite' classes for 'ExcelCall'
# inheritance;
# Import '_format_argument' function for casting non-'ExcelComposite' values
# to 'ExcelArgument' objects;
from .composite import ExcelCompositeType, ExcelComposite, _format_argument


__all__ = ()    # No classes defined in this package should be implicitly
                # imported.


def _repr_arg(arg) -> str:
    # Return string representation of 'arg';
    return arg.function.name if isinstance(arg, ExcelFunctionCall) \
        else repr(arg) if hasattr(arg, "__repr__") \
        else str(arg) if hasattr(arg, "__str__") \
        else arg.__name__ if isinstance(arg, type) \
        else arg.__class__.__name__


class ExcelCallType(ExcelCompositeType):
    """Abstract class for 'ExcelCall' classes."""
    pass


class ExcelCall(ExcelComposite, metaclass=ExcelCallType):
    """
    The 'ExcelCall' subclasses are designed to handle individual calls
    to the 'ExcelFunction' and 'ExcelFormula' classes.

    When one of these two classes are instantiated, it returns an instance
    of the 'ExcelCall' that corresponds with the calling class rather than
    an instance of the calling class.
    
    Arguments passed in this way are stored in the 'ExcelCall' instance,
    which will handle these arguments according to the rules defined in the
    calling 'ExcelFunction' or 'ExcelFormula' subclass.

    ------------------------------------------------------------------

    This class is not intended to be instantiated or subclassed directly.
    Instead, users should subclass either the 'ExcelFunction' or
    'ExcelFormula' calling classes and define the desired handling logic
    within that subclass.

    The user can then instantiate that calling subclass to create an
    'ExcelCall' instance which will follow the logic of that subclass.
    """
    def __init__(self, caller, args: tuple):
        # Instantiate parent classes;
        super().__init__()
        self._caller = caller
        # Cast all arguments to 'ExcelComposite' objects;
        self._arguments = tuple(
            _format_argument(arg) for arg in args
            )
        # Set self as owner of all arguments;
        for arg in self._arguments:
            arg._owner = self

    @property
    def caller(self):
        """Returns call's caller class."""
        return self._caller

    @property
    def name(self) -> str:
        """Returns name of call's caller class."""
        return self._caller.__name__

    @property
    def arguments(self) -> tuple:
        """Returns arguments passed to call's caller class."""
        return self._arguments

    @property
    def required_arguments(self) -> tuple:
        """
        Returns call's arguments that correspond with its caller's required
        arguments.
        """
        req_count = len(self._caller.required_arguments)
        # Return arguments below the number of required arguments;
        return self._arguments[:req_count]

    @property
    def optional_arguments(self) -> tuple:
        """
        Returns call's arguments that correspond with its caller's optional
        arguments.
        """
        req_count = len(self._caller.optional_arguments)
        # Return arguments above the number of required arguments;
        return self._arguments[req_count:]

    def get_value(self):
        """
        Returns the calculated value of the call of based on its arguments.
        """
        return self._caller.get_value(self)

    def get_values(self) -> tuple:
        """
        Returns the calculated value of each of call's arguments based on
        each argument's sub-arguments.
        """
        return (
            arg.get_value() for arg in self._arguments
            )

    def __repr__(self) -> str:
        # Collect required and optional arguments;
        req_args = self.required_arguments
        opt_args = self.optional_arguments
        sep = ", "

        # Join representation strings of each required argument with a comma
        # and space;
        req_string = sep.join(
            _repr_arg(arg) for arg in req_args
            )
        # Join representation strings of each optional argument with a comma
        # and space;
        opt_string = sep.join(
            _repr_arg(arg) for arg in opt_args
            )
        # Enclose optional representation in '[]' if it exists;
        if opt_string or self._caller.is_openended():
            opt_string = f"[{opt_string}]"
        # Join the required and optional argument string with a comma and
        # space;
        arg_string = sep.join((req_string, opt_string))
        # Return the classname of the caller class joined with the
        # required and optional argument string;
        return f"'{self._caller.__name__}' ({arg_string})"


class ExcelFunctionCall(ExcelCall):
    # Set 'ExcelFunctionCall' '__doc__' string to that of 'ExcelCall' for
    # consistency;
    __doc__ = "'ExcelFunction' specific implementation of 'ExcelCall' class" \
              + "\n\n" + ExcelCall.__doc__
    
    @property
    def function(self):
        # 'ExcelFunctionCall' class specific caller property;
        return self._caller
    function.__doc__ = ExcelCall.caller.__doc__
    # Set 'formula' '__doc__' string to that of 'caller' property for
    # consistency;

    def compile(self, indent=None):
        # Join the compile string of call's arguments with a comma and space;
        arg_string = ", ".join(arg.compile() for arg in self._arguments)
        # Join call's name with the joined argument string enclosed in '()';
        return f"{self.name}({arg_string})"


class ExcelFormulaCall(ExcelCall):
    # Set 'ExcelFormulaCall' '__doc__' string to that of 'ExcelCall' for
    # consistency;
    __doc__ = "'ExcelFormula' specific implementation of 'ExcelCall' class" \
              + "\n\n" + ExcelCall.__doc__
    
    @property
    def formula(self):
        # 'ExcelFormulaCall' specific caller property;
        return self._caller
    # Set 'formula' '__doc__' string to that of 'caller' property for c
    # onsistency;
    formula.__doc__ = ExcelCall.caller.__doc__

    def compile(self, indent=None):
        # Join and return compiled strings of call's arguments without a
        # separator;
        formula_string = "".join(
            str(arg) for arg in self._arguments
            )
        return formula_string
