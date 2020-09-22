# ./lib/argument.py

"""
Module defines the 'ExcelArgumentHandler' class for handling argument
parameter definition and enforcement in 'ExceptFunction' and 'ExcelFormula'
classes.

---------------------------------------------------------------------

Users should not import directly from this module.
All relevant classes and constants defined here are imported into the
'lib' package and then into the top-level 'xlformula' package.

Classes and constants can be imported from either package with the
'from [package] import *' statement.

All abstract and meta classes defined here are imported into the 'abc'
module.
"""

from abc import ABCMeta, abstractclassmethod
from collections.abc import Iterable
from itertools import chain


__all__ = (                 # Defines '__all__' for implicit '*' imports;
    )

REQUIRED_ARGUMENTS = "__requiredarguments__"    # Constant for varname that
                                                # identifies the required
                                                # argument parameters for the
                                                # class;
OPTIONAL_ARGUMENTS = "__optionalarguments__"    # Constant for varname that
                                                # identifies the optional
                                                # argument parameters for the
                                                # class;
INHERIT_ARGUMENTS = "__inheritarguments__"      # Constant for varname that
                                                # identifies whether the class
                                                # will inherit the argument
                                                # parameters of its parent
                                                #class;
ARGUMENT_HANDLER_METHOD = "__argumenthandlermethod__"


def _get_funcargs(cls, varname: str) -> tuple:
    """Returns the value of varname in class 'cls' or (...,) by default."""
    return getattr(cls, varname, (...,))

def _funcargs_to_tuple(args: tuple) -> tuple:
    """
    Returns 'arg' cast to a tuple if it is not already a tuple
    e.g. "arg" -> ("arg",)
    """
    if isinstance(args, str) \
    or not isinstance(args, Iterable):
        # Contain 'args' in a tuple if it is not an iterable. Otherwise it
        # converts 'args' to a tuple if it is an iterable;
        # Type checks 'args' against 'str' to ensure it does not split 'args'
        # if it is a string e.g. "arg" -x-> ("a", "r", "g");
        args = (args,)
    elif not isinstance(args, str) \
    and isinstance(args, Iterable):
        args = tuple(args)

    # Convert all 'None' contained in 'args' to '...';
    if None in args:
        args = tuple(
            ... if arg is None else arg for arg in args
            )
    return args     # Return resulting tuple or arguments;

def _inherit_funcargs(varname:str, bases: tuple, args: tuple) -> tuple:
    """
    Chains together all argument parameters defined in each parent class in
    ' bases' and returns it.
    """
    # Collect all required or optional arguments defined in parent classes;
    base_args = (getattr(base, varname, ()) for base in bases)
    return tuple(chain(*base_args, args)) # Return chained argument parameters

def _repr_funcargname(arg) -> str:
    """Returns determined representation string of 'arg'."""
    # Returns "..." instead of 'Ellipsis' as representation string for '...';
    return "..." if arg is ... else str(arg)


class ExcelArgumentHandlerType:
    """Abstract class for 'ExcelArgumentHandler' classes."""
    __inheritarguments__ = False    # Defines class variable that identifies
                                    # if class can inherit argument parameters
                                    # from its parent classes;
    __requiredarguments__ = ()      # Defines class variable that defines the
                                    # argument parameters that a user must
                                    # provide to instantiate the class;
    __optionalarguments__ = ()      # Defines class variable that defines the
                                    # argument parameters that a user can
                                    # provide beyond the required arguments
                                    #  when instantiating the class;
    __argumenthandlermethod__ = ""

    def __new__(cls, name: str, bases: tuple, namespace: dict):
        # Instantiate base class;
        cls = super().__new__(cls, name, bases, namespace)
        # Cast argument parameter variables to tuples;
        req_args = _funcargs_to_tuple(cls.__requiredarguments__)
        opt_args = _funcargs_to_tuple(cls.__optionalarguments__)

        if cls.__inheritarguments__:
            # If class is allowed to inherit argument parameters, collect
            # parent class argument parameters and insert them ahead of
            # arguments defined in this class;
            req_args = _inherit_funcargs(REQUIRED_ARGUMENTS, bases, req_args)
            opt_args = _inherit_funcargs(OPTIONAL_ARGUMENTS, bases, opt_args)

        if cls.__argumenthandlermethod__ \
        and len(req_args) < 1 \
        and len(opt_args) < 1:
            arg_handler = getattr(cls, cls.__argumenthandlermethod__, None)
            if arg_handler \
            and not getattr(arg_handler, "__isabstractmethod__", False):
                if arg_handler.__kwdefaults__:
                    raise TypeError(
                        f"'{arg_handler.__name__}' method for class "\
                        f"'{cls.__name__}' cannot declare keyword-only " \
                        "parameters; All keyword parameters must be " \
                        "positional."
                        )

                arg_names = arg_handler.__code__.co_varnames[1:]
                opt_count = len(arg_handler.__defaults__) \
                            if arg_handler.__defaults__ \
                            else 0
                req_count = len(arg_names) - opt_count

                req_args = _funcargs_to_tuple(arg_names[:req_count])
                opt_args = _funcargs_to_tuple(arg_names[req_count:])

        # Assign processed argument parameters to this class;
        cls.__requiredarguments__ = req_args
        cls.__optionalarguments__ = opt_args

        return cls      # Return this class;

    @property
    def required_arguments(cls) -> tuple:
        """Returns the classes required arguments."""
        return cls.__requiredarguments__

    @property
    def optional_arguments(cls) -> tuple:
        """Returns the classes optional arguments."""
        return cls.__optionalarguments__

    @property
    def arguments(cls) -> tuple:
        """Returns the classes required and optional arguments."""
        # Combine and return class's required and optional arguments;
        return (
            *cls.required_arguments,
            *cls.optional_arguments
            )

    def is_openended(cls) -> bool:
        """
        Return 'True' if an infinite number of arguments can be passed
        to the class;
        """
        # Class is considered 'open-ended' if either its required or optional
        # arguments contain '...';
        return ... in cls.required_arguments \
            or ... in cls.optional_arguments

    def __repr__(cls) -> str:
        # Returns a representation string of class's name and class's
        # required and optional argument parameters joined by a comma and
        # space and enclosed in '()' with optional arguments enclosed in '[]'
        # e.g. ("reqarg1", "reqarg2"), ("optarg1")
        # -> "ArgumentHandlerClass('arg1', 'arg2', ['optarg1'])";
        sep = ", "
        # Join required parameter representation strings with a comma and space;
        req_string = sep.join(
            _repr_funcargname(arg) for arg in cls.required_arguments
            )

        # Collects all but the last paramter in class's optional arguments
        # if the last parameter is '...'. Otherwise, collections all
        # optional argument parameters;
        opt_args = cls.optional_arguments
        endswith_arg = opt_args[-1] if opt_args else None
        endswith_ellipsis = endswith_arg is ...
        opt_args = opt_args[:-1] if endswith_ellipsis else opt_args

        # Joins all collected optional parameters representation strings
        # with a comma and space
        # e.g. ("optarg1", "optarg2") -> "'optarg1', 'optarh2'";
        opt_string = sep.join(
            _repr_funcargname(arg) for arg in opt_args
            )
        # If any optional parameters were collected, they are enclosed in '[]'
        # e.g. "'optarg1', 'optarg2'" -> "['optarg1', 'optarg2']"
        #       "" -> "", "" -x-> "[]";
        if opt_string:
            opt_string = f"[{opt_string}]"
        # Appends "..." with a comman and space to representation string,
        # if the last optional parameter   was '...'
        # e.g. "['optarg1', 'optarg2']" -> "['optarg1', 'optarg2'], ..."
        if endswith_ellipsis:
            endswith_name = _repr_funcargname(endswith_arg)
            comma = ", " if opt_string else ""
            opt_string = comma.join((opt_string, endswith_name))
        # Joins required parameter representation and optional parameters
        # strings with comman and space;
        # e.g. "'reqarg1', 'reqarg2'", " ['optarg1', optarg2], ..."
        # -> "'reqarg1', 'reqarg2', ['optarg1', optarg2], ..."
        arg_string = sep.join(
            string for string in (req_string, opt_string) if string
            )
        # Returns representation string with class name joined with argument
        # parameters enclosed in '()' e.g.
        # "ArgumentHandlerClass('reqarg1', 'reqarg2', ['optarg1', optarg2], ...)";
        return f"{cls.__name__}({arg_string})"


class ExcelArgumentHandler(metaclass=ABCMeta):
    """Base class for handling and enforcing argument parameters."""
    def __new__(cls, *args):
        # Collection required and optional arugment parameters from class;
        req_args = _get_funcargs(cls, REQUIRED_ARGUMENTS)
        opt_args = _get_funcargs(cls, OPTIONAL_ARGUMENTS)
        # Find the number of arguments passed, the number of arguments that
        # are required, and the number of arguments that are optional;
        arg_count = len(args)
        req_count = arg_count if ... in req_args else len(req_args)
        opt_count = arg_count if ... in opt_args else len(opt_args)

        # Format grammatically correct form of "was";
        were = "was" if abs(arg_count) == 1 else "were"
        classname = cls.__class__.__name__
        if req_count and arg_count < req_count:
            # Raise exception if not enough arguments were passed;
            missing = req_args[arg_count:]
            s1 = "" if req_count == 1 else "s"
            s2 = "" if len(missing) < 2 else "s"
            are = "is" if len(missing) < 2 else "are"
            raise TypeError(
                f"'{classname}' requires " \
                f"{req_count} argument{s1}, but " \
                f"{arg_count} {were} given: " \
                f"Argument{s2} {missing} {are} missing."
                )
        elif arg_count > req_count + opt_count:
            # raise exception if too many aruments were passed;
            raise TypeError(
                f"'{classname}' only accepts "\
                f"{req_count} required arguments and " \
                f"{opt_count} optional arguments, but "
                f"{arg_count} {were} given."
                )
        else:
            # Otherwise, instantiate class and return results of
            # '_handle_arguments' method;
            cls = super().__new__(cls)
            return cls._handle_arguments(args)

    @abstractclassmethod
    def _handle_arguments(cls, args: tuple):
        """
        Abstract class method used to handle arguments passed the class
        after arguments have been checked against what is required.
        """
        raise NotImplementedError

    @abstractclassmethod
    def get_value(self):
        """
        Abstract method used to calculate and return the value of contained
        in an 'ExcelCall' object.

        'self' parameter is not an 'ExcelArgumentHandler' class, but instead
        an 'ExcelCall' object.
        """
        raise NotImplementedError
