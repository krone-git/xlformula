# ./abc

"""
Module contains all abstract classes available for type checking.

----------------------------------------------------------------------

It is not recommended for users to inherit from these metaclasses
directly.
"""

from .lib.function import ExcelFunctionType
from .lib.formula import ExcelFormulaType
from .lib.composite import ExcelCompositeType, ExcelStringBuilderType
from .lib.call import ExcelCallType
from .lib.reference import ExcelReferenceType

__all__ = (
    "ExcelFunctionType",
    "ExcelFormulaType",
    "ExcelStringBuilderType",
    "ExcelCompositeType",
    "ExcelCallType",
    "ExcelReferenceType"
    )
