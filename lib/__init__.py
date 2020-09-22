# ./lib

# Import all '*' available imports from 'builtin' package;
from .builtin import *
from .builtin import __all__ as __builtin_all__

# Import all '*' available imports from 'reference' module;
from .reference import *
from .reference import __all__ as __reference_all__

# Import all '*' available imports from 'composite' module;
from .composite import *
from .composite import __all__ as __composite_all__

# Import all '*' available imports from 'formula' module;
from .formula import *
from .formula import __all__ as __formula_all__

# Import all '*' available imports from 'function' module;
# Import 'ExcelFunction' factory classes;
from .function import *
from .function import __all__ as __function_all__
from .function import ExcelFunctionClassFactory, ExcelFunctionCallFactory


# Dynamically generate convenience namespace variables for Excel classes;
_CONVENIENCE_CLASSNAMES = {
    "cell": ExcelCellReference,
    "ref":  ExcelReference,
    "arg":  ExcelArgument,
    "var":  ExcelArgument,
    "func": ExcelFunctionClassFactory,
    "call": ExcelFunctionCallFactory
    }

_vars = vars()                                      # Store namespace 'vars()'
                                                    # to avoid multpile calls
                                                    # to 'vars()';
for _k, _v in _CONVENIENCE_CLASSNAMES.items():
    # EDIT: 9/22/2020 Brandon Krone
    # The original version of this block set varnames to lowercase, uppercase,
    # and titlecase forms of the varname. This has been changed to only set
    # the varname to its lowercase form to avoid confusion and overlap
    # with imported 'ExcelFunction' classes, which are necessarily uppercase.
    # The original block also included single letter varnames. This too has
    # been changed to prevent overlap with other single letter varnames
    # defined by the user after importing this module.
    # If a single letter, convenience varname for any of the contained
    # classes is desired, it should be declared explicitly by the user with
    # the 'as' clause of an 'import' statement;
    
    # Ensure class varname is lowercase, then set it to the appropiate class.
    # Then, include the varname in '_CONVENIENCE_CLASSNAMES' to be included
    # in '__all__' to allow the class to be imported with '*' with the given
    # varname;
    _k = _k.lower()
    _vars.setdefault(_k_, _v)
    _CONVENIENCE_CLASSNAMES.setdefault(_k_, _v)

del _vars, method, methods, _k, _v   # Delete to prevent explicit imports;

# Combine all '*' available imports from subpackages and modules into
# '__all__' of parent package;
# Include generated convenience varnames;
__all__ = (
    *__builtin_all__,
    *__reference_all__,
    *__composite_all__,
    *__formula_all__,
    *__function_all__,
    *_CONVENIENCE_CLASSNAMES
    )

del _CONVENIENCE_CLASSNAMES     # Delete to prevent explicit imports;
