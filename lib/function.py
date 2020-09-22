# ./lib/function.py

"""
Module defines abstract 'ExcelFunction' classes for defining and working
Excel functions.

---------------------------------------------------------------------

Users should not import directly from this module.
All relevant classes and constants defined here are imported into the
'lib' package and then into the top-level 'xlformula' package.

Classes and constants can be imported from either package with the
'from [package] import *' statement.

All abstract and meta classes defined here are imported into the 'abc'
module.
"""

from abc import ABCMeta

from .call import ExcelFunctionCall
# Import 'ExcelArgumentHandlerType' and 'ExcelArgumentHandlerType' classes
# for 'ExcelFunction' class Inheritance;
# Import '__requiredarguments__' and '__optionalarguments__' constants,
# and '_funcargs_to_tuple' function for use in dynamically generated
# 'ExcelFunction' classes;
from .argument import ExcelArgumentHandlerType, ExcelArgumentHandler, \
                        REQUIRED_ARGUMENTS, OPTIONAL_ARGUMENTS, \
                        _funcargs_to_tuple


__all__ = (             # Defines '__all__' for implicit '*' imports;
    "ExcelFunction",    # 'ExcelFunction' class is the only class which
                        # should be implicitly imported.
    )

# Constant for varname that identifies if an 'ExcelFunctionNameHandler'
# subclass is allowed to have a classname that is not uppercase;
IS_BASE_EXCEL_FUNCTION_CLASS = "__isbaseexcelfunctionclass__"
GET_VALUE = "get_value"     # Constant for varname of 'get_value' method
                            # for use in dynamically generating 'ExcelFunction'
                            # classes;

def _notimplemented(self):
    """Raises 'NotImplementedError' exception."""
    # Default function for 'get_value' method when dynamically generating
    # 'ExcelFunction' classes;
    # Raises 'NotImplementedError' to prevent use of 'get_value' method
    # on classes where it is not explicitly defined.
    raise NotImplementedError


class ExcelFunctionNameHandler:
    """
    Absract class for handling and enforcing naming conventions for
    'ExcelFunction' subclasses.
    """
    def __new__(cls, name: str, bases: tuple, namespace: dict):
        # Enforces uppercase naming convention for Excel functions
        # upon creation of 'ExcelFunctionNameHandler' subclasses;

        # Raises 'TypeError' if subclass name is not uppercase and
        # subclass does not contain the '__isbaseexcelfunctionclass__'
        # variable or variable is 'False'.
        # Otherwise, continues with subclass creation.
        if not namespace.pop(IS_BASE_EXCEL_FUNCTION_CLASS, False) \
        and not name.isupper():
            raise TypeError(
                f"Excel function names must be uppercase: " \
                f"'{name}' is not a valid function name."
                )
        else:
            return super().__new__(cls, name, bases, namespace)


class ExcelFunctionArgumentHandlerType(ExcelArgumentHandlerType):
    """
    Abstract class for handling and enforcing required and optional
    arguments for 'ExcelFunction' subclasses.
    """
    def __repr__(cls) -> str:
        # Return shell interface representation of 'ExcelFunction' subclass
        # name combined with representation of 'ExcelFunction' base class;
        return f"<{ExcelFunction.__name__}> {super().__repr__()}"


class ExcelFunctionType(ExcelFunctionArgumentHandlerType,
                        ExcelFunctionNameHandler,
                        ABCMeta):
    """Abstract class for the 'ExcelFunction' class."""
    pass


class ExcelFunctionArgumentHandler(ExcelArgumentHandler):
    __doc__ = ExcelArgumentHandlerType.__doc__
    @classmethod
    def _handle_arguments(cls, args: tuple):
        """
        Override method for '_handle_arguments' abstract method defined in
        'ExcelArgumentHandler'.
        """

        # Returns 'ExcelFunctionCall' object with self as '_caller' and 'args';
        return ExcelFunctionCall(cls, args)


class ExcelFunction(ExcelFunctionArgumentHandler,
                    metaclass=ExcelFunctionType):
    # NOTE: 9/20/2020 Brandon Krone
    # How can we specify default values for optional arguments?;
    # Setting '__optionalarguments__' to dict eliminates the use of '...'
    # to allow infinite arguments to be passed;
    """
    Base implementation for working with Excel functions.

    The 'ExcelFunction' base class is not intended to be instantiated directly.
    Instead, users should define subclasses which inherit from 'ExcelFunction'.

    Instantiating an 'ExcelFunction' subclass returns an 'ExcelFunctionCall'
    object with an association with the calling 'ExcelFunction' subclass.

    Calling the overridden abstract methods defined in the 'ExcelFunction'
    subclass from the 'ExcelFunctionCall' object will perform those operations
    on the 'ExcelFunctionCall' object instead of the 'ExcelFunction' subclass.

    -----------------------------------------------------------------

    To add argument parameters to the 'ExcelFunction' subclass, the user
    must override the '__requiredarguments__' and '__optionalarguments__'
    class variables based on the desired parameters. These variables should
    be tuples of parameters. Both variables default to empty tuples.

    'ExcelFunction' subclasses will raise an exception when called with the
    incorrect amount of arguments. The required and optional argument
    variables can be set to 'None', '...', or to a tuple that contains either
    'None' or '...' to allow an infinite number of arguments to be passed.
    If neither argument variable contains '...' or 'None', the subclass
    will be constrained by the number of defined argument parameters.

    - A call that contains fewer than the required number of arguments will
        always throw an exception.
    - A call that contains more than the required number of arguments will
        not throw an exception, so long as the number of extra arguments
        passed does not surpass the number of specified optional arguments.
    - A subclass that defines no required arguments, but contains optional
        arguments will accept 0 or more arguments up to the number of
        defined optional arguments. However, if no optional arguments are
        defined, it will raise an exception if any arguments are passed.

    All defined argument parameters are positional-only.

    When defining a subclass of an 'ExcelFunction' subclass, the user may
    override the '__inheritarguments__' and set it to 'True' to inherit the
    required and optional argument parameters of the parent 'ExcelFunction'
    subclass.
    """
    __isbaseexcelfunctionclass__ = True


class ExcelFunctionClassFactory:
    """
    Factory class for dynamically generating simple 'ExcelFunction' classes.
    This class is intended only to allow users to felixibly create low-impact
    'ExcelFunction' classes where one does not already exist.

    Abuse of this factory class can cause issues and raise unecessary
    exceptions when defining and compiling an Excel formula.
    """
    def __new__(cls, name: str, *, bases: tuple=(), required: tuple=(),
                optional: tuple=(), get=None, **kwargs):
        name = str(name).upper()                # Cast 'name' to uppercase 'str';
        get = get if get else _notimplemented   # Set 'get' method to raise
                                                # NotImplementedError when
                                                # called by default;

        if isinstance(required, int):
            # Dynamically generate a series of default argument names
            # for both required and optional parameters up to the number
            # specified by 'required' and 'optional' (arg1, arg2, arg3, ...);
            required = ("arg" + str(i) for i in range(required))
        if isinstance(optional, int):
            optional = ("arg" + str(i) for i in range(optional))

        # Append arguments contained in 'required' and 'optional' to any
        # contained in keyword arguments under the keys '__requiredarguments__'
        # and '__optionalarguments__';
        kwargs[REQUIRED_ARGUMENTS] = (
            *kwargs.pop(req_name, ()), *_funcargs_to_tuple(required)
            )
        kwargs[OPTIONAL_ARGUMENTS] = (
            *kwargs.pop(opt_name, ()), *_funcargs_to_tuple(optional)
            )

        # Set 'get_value' method to method provided with keyword 'get' if
        # a method was not not already specified in keyword arguments;
        kwargs.setdefault(GET_VALUE, get)

        # Create and return class;
        cls = type(name, (ExcelFunction, *bases), kwargs)
        return cls


class ExcelFunctionCallFactory:
    """
    Factory class for dynamically generating and instantiating simple
    'ExcelFunction' classes.
    This class is intended only to allow users to felixibly create low-impact
    'ExcelFunction' classes where one does not already exist.

    Abuse of this factory class can cause issues and raise unecessary
    exceptions when defining and compiling an Excel formula.
    """
    def __new__(cls, name: str, *args, bases: tuple=(), get=None):
        # Dynamicall generate 'ExcelFunction' subclass;
        func_cls = ExcelFunctionClassFactory(
            name,
            bases=bases,
            required=args,
            get=get
            )
        # Instantiate generated 'ExcelFunction' subclass and return the
        # resulting 'ExcelFunctionCall' object;
        return func_cls(*args)
