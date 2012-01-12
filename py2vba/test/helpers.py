"""
Test helpers.

"""
from win32com.client import Dispatch

from py2vba import vbast, convert

from excelbt.vbproject import Module, VBProject
from excelbt.imports import import_vbproject
from excelbt.vbide import SCRIPTING_REFERENCE

def _create_function_test_stub(vbfunction):
    dictParameter = vbast.Parameter(vbast.SimpleNameExpression('ResultDict'), vbast.Dictionary)

    statements = []
    if vbfunction.rettype.is_object_type():
        assign_statement_type = vbast.SetStatement
    else:
        assign_statement_type = vbast.LetStatement

    statements.append(assign_statement_type(
        vbast.IndexExpression(
            dictParameter.name,
            [vbast.StringLiteral('ReturnValue')]),
        vbast.IndexExpression(
            vbast.SimpleNameExpression(vbfunction.name),
            [p.name for p in vbfunction.parameters])
        ),
    )
    return vbast.Subroutine(
            vbfunction.name + '_TestStub',
            [dictParameter] + vbfunction.parameters,
            statements)

def lift_vba_function(xl, workbook, ast, fname):
    assert fname in ast.function_namespace

    # Create a code stub that handles proxying of VBA return
    # value into Python.
    test_stub = _create_function_test_stub(ast.function_namespace[fname])

    # Build and import a VBA Project into the workbook.
    print '\n'.join(ast.as_code() + test_stub.as_code())
    module = Module('TestModule', '\n'.join(ast.as_code() + 
                                            test_stub.as_code()))
    project = VBProject()
    project.add_module(module)
    project.add_reference(*SCRIPTING_REFERENCE)

    import_vbproject(workbook, project)

    def lifted(*args):
        resultdict = Dispatch('Scripting.Dictionary')
        xl.Run(test_stub.name, resultdict, *args)
        return resultdict['ReturnValue']
    return lifted

def lift_python_function(code, fname, context):
    env = dict()
    exec code in context, env

    if fname not in env:
        raise StandardError('Function named "%s" not found in code "%s".' % (fname, code))

    return env[fname]

def vbast_from_pycode(code):
    pyast = convert.build_ast_from_code(code)
    walker = convert.PythonASTWalker()
    return walker.walk(pyast)
