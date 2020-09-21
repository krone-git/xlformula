# ./lib

from .builtin import *
from .reference import *
from .composite import *
from .function import *
from .function import ExcelFunctionClassFactory, ExcelFunctionCallFactory

# Set convenience class variables;
n = N = ref = Ref = ExcelReference
r = R = Range = ExcelRangeReference
arg = Arg = v = V = var = Var = ExcelArgument
func = FUNC = ExcelFunctionClassFactory
f = F = ExcelFunctionCallFactory

del ExcelFunctionClassFactory, ExcelFunctionCallFactory # Delete to limit
                                                        # direct access to
                                                        # these classes;
