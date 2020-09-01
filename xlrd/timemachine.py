##
# <p>Copyright (c) 2006-2012 Stephen John Machin, Lingfo Pty Ltd</p>
# <p>This module is part of the xlrd package, which is released under a BSD-style licence.</p>
##


def fprintf(f, fmt, *vargs):
    fmt = fmt.replace("%r", "%a")
    if fmt.endswith("\n"):
        print(fmt[:-1] % vargs, file=f)
    else:
        print(fmt % vargs, end=" ", file=f)
