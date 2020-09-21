# ./lib

from .builtin import *
from .builtin import __all__ as __builtin_all__

from .reference import *
from .reference import __all__ as __reference_all__

from .composite import *
from .composite import __all__ as __composite_all__

from .formula import *
from .formula import __all__ as __formula_all__

from .function import *
from .function import __all__ as __function_all__
from .function import ExcelFunctionClassFactory, ExcelFunctionCallFactory


_CONVENIENCE_CLASSNAMES = {
    "n":    ExcelReference,
    "ref":  ExcelReference,
    "arg":  ExcelArgument,
    "v":    ExcelArgument,
    "var":  ExcelArgument,
    "f":    ExcelFunctionCallFactory,
    "func": ExcelFunctionClassFactory,
    }

_vars = vars()
for _k, _v in _CONVENIENCE_CLASSNAMES.items():
    methods = (str.lower, str.title, str.upper) \
              if len(_k) > 1 else (str.lower, str.upper)
    for method in methods:
        _vars[method(_k)] = _v

del _vars, method, methods, _k, _v   # Delete to prevent
                                                            # explicit imports;

__all__ = (
    *__builtin_all__,
    *__reference_all__,
    *__composite_all__,
    *__formula_all__,
    *__function_all__,
    *_CONVENIENCE_CLASSNAMES
    )

del _CONVENIENCE_CLASSNAMES

del ExcelFunctionClassFactory, ExcelFunctionCallFactory     # Delete to limit
                                                            # direct access to
                                                            # these classes;
