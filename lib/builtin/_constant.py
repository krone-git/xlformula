# ./lib/builtin/constants.py

"""
"""


_CONSTANTS = {
    "TRUE": True,
    "FALSE": False
    }

_vars = vars()
for k, v in _CONSTANTS.items():
    k = k.upper()
    _vars.setdefault(k, v)
    _CONSTANTS[k] = v

del _vars, k, v

__all__ = (
    *_CONSTANTS.keys(),
    )

del _CONSTANTS
