
from lib import *

print(ref)
filterstr = BuildFilterString(TRUE, ref("A2"), ref("B1"), ":")
print(filterstr.compile())
