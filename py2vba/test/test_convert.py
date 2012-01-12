from py2vba import vbast
from py2vba.convert import vbmeta
from py2vba.vbast import Integer

from helpers import vbast_from_pycode, lift_vba_function, lift_python_function

def test_basic_function(xl, workbook):
    CODE = '''
@vbmeta(x=Integer, y=Integer, rettype=Integer)
def add(x, y):
    return x + y'''

    ast = vbast_from_pycode(CODE)
    vbafcn = lift_vba_function(xl, workbook, ast, 'add')
    pyfcn = lift_python_function(CODE, 'add', globals())

    pyresult = pyfcn(5,4)
    vbaresult = vbafcn(5,4)

    assert pyresult == vbaresult

def test_array_literal(xl, workbook):
    CODE = '''
@vbmeta(rettype=Integer)
def arraytest():
    a = [1,2,3,4.4]
    return a[0] + a[2]
'''

    ast = vbast_from_pycode(CODE)
    vbafcn = lift_vba_function(xl, workbook, ast, 'arraytest')
    pyfcn = lift_python_function(CODE, 'arraytest', globals())

    pyresult = pyfcn()
    vbaresult = vbafcn()

    assert pyresult == vbaresult

def test_dictionary_literal(xl, workbook):
    CODE = '''
@vbmeta(rettype=Integer)
def dicttest():
    d = {'hello' : 1, 'world' : 2}
    return d['hello']
'''
    ast = vbast_from_pycode(CODE)
    vbafcn = lift_vba_function(xl, workbook, ast, 'dicttest')
    pyfcn = lift_python_function(CODE, 'dicttest', globals())

    pyresult = pyfcn()
    vbaresult = vbafcn()

    assert pyresult == vbaresult
