from abc import ABCMeta, abstractmethod


class ExcelFunctionType(ABCMeta):
    def __new__(cls, name, bases, namespace):
        if not namespace.pop("__isbaseexcelfunctionclass__", False) \
        and not name.isupper():
            raise TypeError(
                f"<{cls.__name__}> class name must be uppercase: " \
                f"'{name}' is invalid."
                )
        else:
            return super().__new__(cls, name, bases, namespace)


class ExcelFunctionArgumentHandler:
    __requiredarguments__ = ()
    __optionalarguments__ = ()

    def __init__(self, *args):
        req_args = self.__requiredarguments__
        opt_args = self.__optionalarguments__
        
        arg_count = len(args)
        req_count = 0 if req_args is None else len(req_args)
        opt_count = arg_count if not req_count or opt_args is None \
                    else len(opt_args)
            
        were = "was" if abs(arg_count) == 1 else "were"

        if req_count and arg_count < req_count:
            missing = self.__requiredarguments__[arg_count:]
            s1 = "" if req_count == 1 else "s"
            s2 = "" if len(missing) < 2 else "s"
            are = "is" if len(missing) < 2 else "are"
            raise TypeError(
                f"'{self.__class__.__name__}' requires " \
                f"{req_count} argument{s1}, but " \
                f"{arg_count} {were} given: " \
                f"Argument{s2} {missing} {are} missing."
                )
        elif arg_count > req_count + opt_count:
            raise TypeError(
                f"'{self.__class__.__name__}' only accepts "\
                f"{req_count} required arguments and " \
                f"{opt_count} optional arguments, but "
                f"{arg_count} {were} given."
                )
        else:    
            self._arguments = tuple(
                arg if isinstance(arg, ExcelFormulaStringComponent) \
                else ExcelArgumentValue(arg) for arg in args
                )

    @classmethod
    def required_arguments(cls):
        return cls.__requiredarguments__ if cls.__requiredarguments__ else ()

    @classmethod
    def optional_arguments(cls):
        return cls.__optionalarguments__ if cls.__optionalarguments__ else ()

    @classmethod
    def arguments(cls):
        return (
            *cls.required_arguments(),
            *cls.optional_arguments()
            )

    @classmethod
    def is_openended(cls):
        return cls.__optionalarguments__ is None


##class InheritableArgumentHandler(ExcelFunctionArgumentHandler):
##    def __new__(cls):
##        req_args = cls.__requiredarguments__
##        opt_args = cls.__optionalarguments__
##        cls = super().__new__(cls)
##        cls.__requiredarguments__ = (
##            *req_args,
##            *cls.__requiredarguments__
##            )
##        cls.__optionalarguments__ = (
##            *opt_args,
##            cls.__optionalarguments__
##            )
##        return cls


class ExcelFormulaStringComponent(metaclass=ABCMeta):
    def __new__(cls, *args, **kwargs):
        cls.__str__ = cls.compile
        return super().__new__(cls)
    
    @abstractmethod
    def compile(self, *args, **kwargs):
        raise NotImplementedError


class ExcelArgumentValue(ExcelFormulaStringComponent):
    def __init__(self, value):
        self._value = value

    @property
    def value(self):
        return self._value

    def compile(self):
        return '""' if self._value is None or self._value == "" \
               else str(self._value)


class ExcelFunction(ExcelFunctionArgumentHandler,
                    ExcelFormulaStringComponent,
                    metaclass=ExcelFunctionType
                    ):
    __isbaseexcelfunctionclass__ = True

    def compile(self, *args, **kwargs):
        arg_string = ", ".join(arg.compile() for arg in self._arguments)
        return f"{self.__class__.__name__}({arg_string})"

    def __repr__(self):
        req_count = len(self.required arguments())
        req_args = self._arguments[:req_count]
        opt_args = self._arguments[req_count:]
        
        req_string = ", ".join(
            f"'{arg.__class__.__name__}'" for arg in req_args
            )
        opt_string = ", ".join(
            f"'{arg.__class__.__name__}'" for arg in opt_args
            )
        if opt_string or self.is_openended():
            opt_string = f"[{opt_string}]"
        
        arg_string = ", ".join(req_string, opt_string)
        return f"'{self.__class__.__name__}' ({arg_string})"

    def required_values(self):
        req_count = len(self.required arguments())
        return req_args = self._arguments[:req_count]

    def optional_values(self):
        req_count = len(self.required arguments())
        return req_args = self._arguments[req_count:]


## ExcelFunction class needs to have instance properties and methods
## split out into a separate, ExcelFunctionCall class.
##
## ExcelFunction should handle creating and supplying values to
## instances of ExcelFunctionCall.
##    
## ExcelFunctionCall instance should hold individual argument values
## and handle formula string building.
## ExcelFunctionCall should reference back to parent class for
## values such as 'function name', 'required arguments' and
## 'optional arguments'
class ExcelFunctionCall:
    pass


class ExcelFormula(ExcelFormulaStringComponent):
    pass


class ExcelFunctionClassFactory:
    def __new__(cls, name, *, bases=(), required=(), optional=(), **kwargs):
        req_name, opt_name = "__requiredarguments__", "__optionalarguments__"

        if isinstance(required, int):
            required = ("arg" + str(i) for i in range(required))
        if isinstance(optional, int):
            optional = ("arg" + str(i) for i in range(optional))
        
        kwargs[req_name] = None if required is None \
                           else (*kwargs.pop(req_name, ()), *required)
        kwargs[opt_name] = None if optional is None \
                           else (*kwargs.pop(opt_name, ()), *optional)

        cls = type(name, (ExcelFunction, *bases), kwargs)
        return cls


class ExcelFunctionFactory:
    def __new__(cls, name, *args, bases=()):
        func_cls = ExcelFunctionBuilder(name, bases=bases, required=args)
        return func_cls(*args)


Var = ExcelArgumentValue

FUNC = ExcelFunctionClassFactory

f = F = ExcelFunctionFactory



myfunc = FUNC("MYFUNC", required=2, optional=None)
inst = myfunc(*range(5))

print(inst)
print(inst.required_arguments())
