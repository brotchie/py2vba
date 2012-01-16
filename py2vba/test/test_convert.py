import py.test

from py2vba import vbast
from py2vba.convert import vbmeta
from py2vba.vbast import Integer, String, Collection

from helpers import lift_code_to_py_and_vba_functions

def test_basic_function(xl, workbook):
    CODE = '''
@vbmeta(x=Integer, y=Integer, rettype=Integer)
def add(x, y):
    return x + y'''

    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'add', globals(), xl, workbook)

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

    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'arraytest', globals(), xl, workbook)

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
    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'dicttest', globals(), xl, workbook)

    pyresult = pyfcn()
    vbaresult = vbafcn()

    assert pyresult == vbaresult

def test_function_call(xl, workbook):
    CODE = '''
@vbmeta(x=Integer, y=Integer, rettype=Integer)
def add(x, y):
    return x + y

@vbmeta(x=Integer, y=Integer, rettype=Integer)
def sumsquare(x, y):
    return add(x*x, y*y)
'''
    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'sumsquare', globals(), xl, workbook)

    pyresult = pyfcn(5,4)
    vbaresult = vbafcn(5,4)

    assert pyresult == vbaresult

def test_array_and_dict_assignments(xl, workbook):
    CODE = '''
@vbmeta(rettype=Integer)
def add(x, y):
    return x + y

@vbmeta(rettype=Integer)
def sumsquare(x, y):
    return add(x*x, y*y)

@vbmeta(rettype=Integer)
def complicated_assignments():
    x = [{'name' : 'James', 'age' : 27},
         {'name' : 'Bob', 'age' : 22}]
    return sumsquare(x[0]['age'], x[1]['age'])

'''
    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'complicated_assignments', globals(), xl, workbook)

    pyresult = pyfcn()
    vbaresult = vbafcn()

    assert pyresult == vbaresult

#def test_integer_for_loop(xl, workbook):
#    CODE = '''
#@vbmeta(rettype=Integer)
#def factorial(n):
#    result = 1
#    for i in range(2,n+1):
#        result *= i
#    return result'''
#
#    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'factorial', globals(), xl, workbook)
#
#    pyresult = pyfcn(5)
#    vbaresult = vbafcn(5)
#
#    assert pyresult == vbaresult

def test_basic_class(xl, workbook):
    CODE = '''
class TestClass(object):
    @vbmeta(name=String)
    def __init__(self, name):
        self.name = name
@vbmeta(rettype=String)
def test():
    t = TestClass('James')
    return t.name
'''
    
    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'test', globals(), xl, workbook)

    pyresult = pyfcn()
    vbaresult = vbafcn()

    assert pyresult == vbaresult

def test_class_relationship(xl, workbook):
    CODE = '''
class Person(object):
    @vbmeta(name=String, age=Integer)
    def __init__(self, name, age):
        self.name = name
        self.age = age

class Company(object):
    @vbmeta(employees=Collection)
    def __init__(self, employees):
        self.employees = employees

@vbmeta(rettype=String)
def test():
    c = Company([
            Person("James", 27),
            Person("Bob", 22)
        ])
    return c.employees[1].name
'''
    pyfcn, vbafcn = lift_code_to_py_and_vba_functions(CODE, 'test', globals(), xl, workbook)

    pyresult = pyfcn()
    vbaresult = vbafcn()

    assert pyresult == vbaresult


