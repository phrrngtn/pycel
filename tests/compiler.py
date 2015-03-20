from pycel.compiler import ExcelCompiler


def test_compiler():
    ec = ExcelCompiler('H:/work//ExcelPythonExporter/PythonExpressionExample2.xlsx')
    code = ec.named_range2code('output')
    print code
    assert len(code) > 0
