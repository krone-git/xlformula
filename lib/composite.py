from abc import ABCMeta, abstractmethod
from functools import wraps

from .clipboard import _copy_to_clipboard


__all__ = (
    "ExcelArgument",
    )

_EXPRESSION_OPERATORS = {
    # Logical
    "__lt__": "<",
    "__le__": "<=",
    "__eq__": "=",
    "__ne__": "<>",
    "__gt__": ">",
    "__ge__": ">=",

    # Arithmetic
    "__add__": "+",
    "__sub__": "-",
    "__mul__": "*",
    "__div__": "/",
    "__pow__": "^",

    # String
    "__and__": "&",

    # "": "",
    }

_OPERATOR_OVERRIDE_METHODS = {
    "__add__": (lambda former, latter: "".join((former, latter)))
    }


def _format_argument(arg):
    if isinstance(arg, tuple) and len(arg) == 1:
        arg = _to_composite(arg)
        arg._priority = True
    else:
        arg = _to_composite(arg)

    if not isinstance(arg, ExcelComposite):
        raise TypeError(
            f"Type {type(arg)} is not a recognized datatype."
            )

    return arg

def _to_composite(arg):
    return arg if isinstance(arg, ExcelComposite) \
        else ExcelArgument(arg)

def _construct_expression(operator):
    func = (
        lambda self, other: ExcelExpression(self, other, operator)
        )
    return func

def _finalize_formula_string(func):
    @wraps(func)
    def _finalize(self):
        formula_string = func(self)
        if self.is_master():
            formula_string = f"={formula_string}"
        return formula_string
    return _finalize

def _prioritize_formula_string(func):
    @wraps(func)
    def _prioritize(self):
        formula_string = func(self)
        if self.has_priority():
            formula_string = f"({formula_string})"
        return formula_string
    return _prioritize

def _compilemethod(func):
    @wraps(func)
    def _compile(self):
        method = getattr(self, func.__name__)
        return method()
    return _compile


class ExcelStringBuilderType(ABCMeta):
    def __new__(cls, name, bases, namespace):
        cls = super().__new__(cls, name, bases, namespace)
        if hasattr(cls, "compile"):
            cls.__str__ = cls.to_string = _compilemethod(cls.compile)
        return cls


class ExcelStringBuilder(metaclass=ExcelStringBuilderType):
    @abstractmethod
    def compile(self, indent=None):
        raise NotImplementedError

    def to_clipboard(self, indent=None):
        string = self.compile(indent=indent)
        _copy_to_clipboard(string)
        return self


class ExcelCompositeType(ExcelStringBuilderType, ABCMeta):
    def __new__(cls, name, bases, namespace):
        cls = super().__new__(cls, name, bases, namespace)
        if hasattr(cls, "compile"):
            cls.compile = _prioritize_formula_string(cls.compile)
            cls.compile = _finalize_formula_string(cls.compile)
        return cls


class ExcelComposite(ExcelStringBuilder, metaclass=ExcelCompositeType):
    _vars = vars()
    for k in _EXPRESSION_OPERATORS:
        k = k.lower()
        _vars[k] = _construct_expression(k)

    del _vars, k

    def __init__(self):
        self._owner = None
        self._priority = False

    @property
    def master(self):
        return self if self._owner is None else self._owner.master

    def is_master(self):
        return self._owner is None

    def has_priority(self):
        return self._priority

    @abstractmethod
    def get_value(self):
        raise NotImplementedError


class ExcelArgument(ExcelComposite):
    def __init__(self, value):
        super().__init__()
        self._value = value

    def get_value(self):
        return self._value

    def compile(self, indent=None):
        return '""' if self._value is None or self._value == "" \
            else str(self._value).upper() if isinstance(self._value, bool) \
            else f'"{self._value}"' if isinstance(self._value, str) \
            else str(self._value)

    def __repr__(self):
        return f"Arg {self.compile()}"


class ExcelExpression(ExcelComposite):
    def __init__(self, former, latter, operator):
        super().__init__()
        self._former = _format_argument(former)
        self._former._owner = self
        self._latter = _format_argument(latter)
        self._latter._owner = self
        self._operator = operator

    @property
    def operator(self):
        return _EXPRESSION_OPERATORS[self._operator]

    def get_value(self):
        former_value = self._former.get_value(),
        latter_value = self._latter.get_value()

        if self._operator in _OPERATOR_OVERRIDE_METHODS:
            operator_method = _OPERATOR_OVERRIDE_METHODS[self._operator]
        else:
            operator_method = getattr(former_value, self._operator)

        return operator_method(former_value, latter_value)

    def compile(self, indent=None):
        return "".join((
            str(self._former),
            self.operator,
            str(self._latter)
            ))
        return string

    def __repr__(self):
        repr_string = "".join((
            str(self._former),
            self.operator,
            str(self._latter)
            ))
        return f"Expression ({repr_string})"
