
# We use this to read .xlsx workbooks as the support for named ranges
# is better than the stuff I hacked into xlrd (and -- most importantly
# -- it is independently packaged)
import openpyxl

from openpyxl.exceptions import (
    CellCoordinatesException,
)

from networkx.classes.digraph import DiGraph

import logging
from logging import debug

#
# This stuff does not have a setup.py/pypi yet. I will do it and publish on our
# internal pypi and send a pull request to the author on github.
#
from pycel.excelcompiler import (
    shunting_yard,
    RangeNode,
    OperatorNode,
    FunctionNode,
    build_ast,
)


logging.getLogger().setLevel('DEBUG')
ec = ExcelCompiler('PythonExpressionExample2.xlsx')
print "remapped pycel compiler output: %s" % (ec.named_range2code('output'))

nr = 'output'
wb = openpyxl.load_workbook('PythonExpressionExample2.xlsx')
ec1 = ExcelCompiler(workbook=wb)
print "named range '%s' : %s" % (nr, ec.named_range2code(nr))

