"""
Test helpers.

"""
from win32com.client import Dispatch

from py2vba import vbast, convert, export

from excelbt.vbproject import Module, VBProject, ClassModule
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
    ast.code.append(test_stub)

    project = VBProject()
    project.add_reference(*SCRIPTING_REFERENCE)
    export.add_procedural_module_to_vbproject(project, ast)
    import_vbproject(workbook, project)

    def lifted(*args):
        resultdict = Dispatch('Scripting.Dictionary')
        xl.Run(test_stub.name, resultdict, *args)
        return resultdict['ReturnValue']
    return lifted

def lift_python_function(code, fname, context):
    # Make a copy of context so we don't accidently
    # modify globals().
    context = dict(context)
    env = dict()

    exec code in context, env

    if fname not in env:
        raise StandardError('Function named "%s" not found in code "%s".' % (fname, code))

    def lifted(*args):
        context.update(env)
        env['args'] = args
        exec 'result = %s(*args)' % (fname,) in context, env
        return env['result']

    return lifted

def vbast_from_pycode(code):
    pyast = convert.build_ast_from_code(code)
    walker = convert.PythonASTWalker()
    return walker.walk(pyast)

def lift_code_to_py_and_vba_functions(CODE, fname, pyenviron, xl, workbook):
    ast = vbast_from_pycode(CODE)
    pyfcn = lift_python_function(CODE, fname, pyenviron)
    vbafcn = lift_vba_function(xl, workbook, ast, fname)

    return pyfcn, vbafcn
