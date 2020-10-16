# ./lib/composite.py

"""
Module defines base, composite pattern classes, string builder classes,
argument value handling classes and operator expression handling classes for
constructing formula strings.

---------------------------------------------------------------------

Users should not import directly from this module.
All relevant classes and constants defined here are imported into the
'lib' package and then into the top-level 'xlformula' package.

Classes and constants can be imported from either package with the
'from [package] import *' statement.

All abstract and meta classes defined here are imported into the 'abc'
module.
"""

from abc import ABCMeta, abstractmethod
from functools import wraps


from .clipboard import _copy_to_clipboard   # Import '_copy_to_clipboard'
                                            # function for convenience when
                                            # moving compiled formula strings
                                            # from api to Excel;


__all__ = (             # Defines '__all__' for implicit '*' imports;
    "ExcelArgument",    # Only 'ExcelArgument' class should be implicitly
                        # imported from this module;
    )

_EXPRESSION_OPERATORS = {
    # Define operators and dunder operator methods that correspond to those
    # found in Excel. These methods will be used to record expressions whose
    # text need to be included in the compiled formula string. If these
    # expressions where not recorded, the formula string will show these
    # expressions post-calculation.
    # e.g. 1 + 2 -> "1 + 2" (desired), 1 + 2 -x-> "3" (undesired);

    # "": "",           # Expression dunder varname: Expression

    # Logical operators;
    "__lt__": "<",          # 'Less than' operator;
    "__le__": "<=",         # 'Less than or equal to' operator;
    "__eq__": "=",          # 'Equal to' operator
                            # (= in Excel rather than == in Python);
    "__ne__": "<>",         # 'Not equal to'operator
                            # (<> in Excel rather than != in Python);
    "__gt__": ">",          # 'Greater than' operator;
    "__ge__": ">=",         # 'Greater than or equal to' operator;

    # Arithmetic operators;
    "__add__": "+",         # Addition arithmetic operator;
    "__sub__": "-",         # Subtraction arithmetic operator;
    "__mul__": "*",         # Multiplication arithmetic operator;
    "__div__": "/",         # Division arithmetic operator;
    "__pow__": "^",         # Exponent arithmetic operator
                            # (^ in Excel rather than ** in Python);

    # String operators;
    "__and__": "&",         # String concatenation operator.
                            # The '&' operator is a bitwise operation in
                            # Python. However, it's a string concatenation
                            # operator in Excel;
    }

# Overrides the dunder method for '&' operator from bitwise operation to
# concatenation;
_OPERATOR_OVERRIDE_METHODS = {
    "__and__": (lambda former, latter: "".join((former, latter)))
    }


def _format_argument(arg):
    """
    Type checks and formats 'arg' by casting it to an 'ExcelComposite' object.
    """
    if isinstance(arg, tuple) and len(arg) == 1:
        # If the argument is contained with in a tuple, it sets the
        # expression resolution priority above the default resolution order;
        # Casts the argument to an 'ExcelComposite' object and sets it's
        # priority to 'True'.
        # Otherwise it only casts the arg to an 'ExcelComposite' object;
        arg = _to_component(arg)
        arg._priority = True
    else:
        arg = _to_component(arg)

    if not isinstance(arg, ExcelComponent):
        # If 'arg' cannot be cast to an 'ExcelComposite' object, raise
        # TypeError;

        # NOTE: 9/21/2020 Brandon Krone
        # Incompatible datatype error handling needs to be moved to another
        # function;
        raise TypeError(
            f"Type {type(arg)} is not a recognized datatype."
            )

    return arg      # Return the formatted argument;

def _to_component(arg):
    """
    Casts 'arg' to an 'ExcelArgument' object if it is not already an
    'ExcelComposite' object.
    """
    return arg if isinstance(arg, ExcelComponent) \
        else ExcelArgument(arg)

def _construct_expression_method(operator):
    """
    Dynamically generates an 'ExcelExpression' instantiation method with
    a given operator.
    """
    func = (
        lambda self, other: ExcelExpression(self, other, operator)
        )
    return func

def _finalize_formula_string(func):
    """
    Decorator function for the 'compile' method that appends '=' to the
    beginning of the compiled formula string of the top-level formula.
    """
    @wraps(func)
    def _finalize(self):                # Wrap function;
        formula_string = func(self)     # Get initial formula string;
        if self.is_master():
            # If the object is the top-level object, append '=' to the string;
            formula_string = f"={formula_string}"
        return formula_string
    return _finalize                    # Return wrapped function;

def _prioritize_formula_string(func):
    """
    Decorator function for the 'compile' method that appends '()'s the the
    head and tail of a compiled formula to ensure 'order of operation'
    priority.
    e.g.
        (2 + 3) * 2     --->    "(2 + 3) * 2" = 10
        2 + 3 * 2       -x->    "2 + 3 * 2" = 8
    """
    @wraps(func)
    def _prioritize(self):              # Wrap function;
        formula_string = func(self)     # Get initial formula string;
        if self.has_priority():
            # If object._priority is 'True', append '()' to head and tail
            # of the string;
            formula_string = f"({formula_string})"
        return formula_string
    return _prioritize                  # Return wrapped function;

def _compilemethod(func):
    """
    Decorator method for the 'compile' method used when setting the
    '__str__' and 'to_string' methods of the class. This decorator ensures
    that the version of the 'compile' method set to those varnames
    is the the 'compile' method of the top level class instead of that of
    the parent class.
    """
    @wraps(func)            # Wrap function;
    def _compile(self):
        # Return the top level version of the 'compile' method;
        method = getattr(self, func.__name__)
        return method()
    return _compile         # Return wrapped function;


class ExcelStringBuilderType(ABCMeta):
    """Abstract class for 'ExcelStringBuilder' classes."""
    def __new__(cls, name, bases, namespace):
        cls = super().__new__(cls, name, bases, namespace)
        if hasattr(cls, "compile"):
            # If the class defines the 'compile' method, set varnames
            # '__str__' and 'to_string' to the top-level version of the
            # 'compile' method;
            cls.__str__ = cls.to_string = _compilemethod(cls.compile)
        return cls


class ExcelStringBuilder(metaclass=ExcelStringBuilderType):
    """
    Base class for objects that contribute to the compiling of formula strings.

    -----------------------------------------------------------------

    This class is not intended to be subclassed.
    """
    @abstractmethod
    def compile(self, indent=None):
        # Defines abstract method 'compile' for generating formula strings;
        raise NotImplementedError

    def to_clipboard(self, indent=None):
        """
        Convenience method that copies the compiled formula string to the
        clipboard for easy pasting of formula string.
        """
        # Compile the formula string then clear the clipboard and append the
        # formula string to the clipboard;
        string = self.compile(indent=indent)
        _copy_to_clipboard(string)
        return self


class ExcelComponent(metaclass=ABCMeta):
    """
    Base class for objects that will be interacted with directly by the user
    when designing and compiling formulas.

    -----------------------------------------------------------------

    This class is not intended to be subclassed.
    """

    @abstractmethod
    def get_value(self):
        """Abstract method that calculates returns the value of the object."""
        raise NotImplementedError


class ExcelCompositeType(ExcelStringBuilderType, ABCMeta):
    """Abstract class for the 'ExcelComposite' classes."""
    def __new__(cls, name, bases, namespace):
        cls = super().__new__(cls, name, bases, namespace)
        if hasattr(cls, "compile"):
            # If the class defines the 'compile' method, wrap the method so
            # that is will append '=' to the head if it is the top-level
            # object while also enclosing the string in '()'s if it has
            # 'order of operations' priority;
            cls.compile = _prioritize_formula_string(cls.compile)
            cls.compile = _finalize_formula_string(cls.compile)
        return cls


class ExcelComposite(ExcelStringBuilder, ExcelComponent,
                        metaclass=ExcelCompositeType):
    """
    Base class for objects that will be interacted with directly by the user
    when designing and compiling formulas.

    -----------------------------------------------------------------

    This class is not intended to be subclassed.
    """
    _vars = vars()                      # Stores 'vars()' to avoid multiple
                                        # calls to 'vars()';
    for k in _EXPRESSION_OPERATORS:
        # Sets each expression operator dunder method to return an
        # 'ExcelExpression' instance when called;
        k = k.lower()                               # Ensures method name is
                                                    # lowercase;
        _vars[k] = _construct_expression_method(k)

    del _vars, k        # Delete to prevent access to these temporary variables;

    def __init__(self):
        # '_owner' and '_priority' variables are set tpo 'None' and 'False'
        # by default. These variables will be set later by an 'owner' object;
        self._owner = None          # The '_owner' variable references the
                                    # parent of object of this object;
        self._priority = False      # Identifies if this object has
                                    # 'order of operations' priority;

    @property
    def master(self):
        """Returns the top-level object in the formula."""
        # If the object does not have an owner, it returns itself as the
        # top-level object;
        return self if self._owner is None else self._owner.master

    def is_master(self):
        """
        Returns 'True' if the object is the top-level object in the formula.
        """
        return self._owner is None

    def has_priority(self):
        """Returns 'True' if the object has 'order of operations' priority."""
        return self._priority


class ExcelArgument(ExcelComposite):
    """
    Implementation for working with base-level function and formula arguments.
    An 'ExcelArgument' instance stores a primitive datatype value for use in
    calculating the value of a function or formula as well as compiling a
    formula string.

    -----------------------------------------------------------------

    This class is not intended to be subclassed.
    """
    def __init__(self, value):
        super().__init__()
        self._value = value     # Set the stored, primitive, value of this
                                # argument;

    def get_value(self):
        """Return the primitive, argument value stored by this object."""
        return self._value

    def compile(self, indent=None):
        # If argument's value is an empty string return '""'. Empty strings
        # in Excel are represented as '""'. Simply returning the empty string
        # would compile to a join of other arguments with out a character
        # in between;
        # If the argument's value is a string, enclose it in '""'s.
        # Otherwise return the base 'str()' representation of the value;
        return '""' if self._value is None or self._value == "" \
            else str(self._value).upper() if isinstance(self._value, bool) \
            else f'"{self._value}"' if isinstance(self._value, str) \
            else str(self._value)

    def __repr__(self):
        # Return the compiled formula string with the 'ExcelArgument'
        # identifier, 'Arg';
        return f"Arg {self.compile()}"


class ExcelExpression(ExcelComposite):
    """
    Implementation for working with logical, arithmetic and string
    manipulation expressions.

    The 'ExcelExpression' class stores expressions with their arguments
    and operators so that they can be included in the compiled formula
    string before the expression is calculated.
    e.g.
        1 + 2   ->  compile()   --->    "1 + 2" (desired)
        1 + 2   ->  compile()   -x->    "3"     (undesired)

    Performing operations on any 'ExcelComposite' class will return an
    'ExcelExpression' object.

    -----------------------------------------------------------------

    This class is not intended to be subclassed.
    """
    def __init__(self, former, latter, operator):
        super().__init__()
        # Assign the expression's first argument object and set its '_owner'
        # to this expression;
        self._former = _format_argument(former)
        self._former._owner = self
        # Assign the expression's second argument object and set its '_owner'
        # to this expression;
        self._latter = _format_argument(latter)
        self._latter._owner = self
        # Assign the expression's operation method's varname;
        self._operator = operator

    @property
    def operator(self):
        """Return the operation character of the expression."""
        return _EXPRESSION_OPERATORS[self._operator]

    def get_value(self):
        """Return the calculated value of the expression."""
        # Get the calculated values of the first and second expression
        # arguments;
        former_value = self._former.get_value(),
        latter_value = self._latter.get_value()

        if self._operator in _OPERATOR_OVERRIDE_METHODS:
            # If the default operator method has been overrideen return the
            # override method. Otherwise, return the default operation method;
            operator_method = _OPERATOR_OVERRIDE_METHODS[self._operator]
        else:
            operator_method = getattr(former_value, self._operator)

        # Call and return the result of the operation method;
        return operator_method(former_value, latter_value)

    def compile(self, indent=None):
        # Join and return the compiled formula strings of the the expression
        # arguments combined with the operation character of the expression;
        return "".join((
            str(self._former),
            self.operator,
            str(self._latter)
            ))
        return string

    def __repr__(self):
        # Join the compiled formula strings of the the expression arguments
        # combined with the operation character of the expression;
        repr_string = "".join((
            str(self._former),
            self.operator,
            str(self._latter)
            ))
        return f"Expression ({repr_string})"    # Return representation string
                                                # with 'Expression' identifier;
