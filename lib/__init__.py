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
    "n":    ExcelReference,
    "ref":  ExcelReference,
    "arg":  ExcelArgument,
    "v":    ExcelArgument,
    "var":  ExcelArgument,
    "f":    ExcelFunctionCallFactory,
    "func": ExcelFunctionClassFactory,
    }

_vars = vars()                                      # Store namespace 'vars()'
                                                    # to avoid multpile calls
                                                    # to 'vars()';
for _k, _v in _CONVENIENCE_CLASSNAMES.items():
    # Generate uppercase, lowercase, and titlecase varnames for each class;
    # Do not pass varname to 'str.title'
    methods = (str.lower, str.title, str.upper) \
              if len(_k) > 1 \
              else (str.lower, str.upper)

    for method in methods:
        # Apply the 'str' method to the varname and set it to the appropriate
        # class;
        _k_ = method(_k)
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

del ExcelFunctionClassFactory, \
    ExcelFunctionCallFactory, \
    _CONVENIENCE_CLASSNAMES     # Delete to limit direct access to these
                                # classes/variables;
