Python to VBA Code Converter
============================
py2vba converts a sub-set of Python into VBA code suitable for use as macros
within Excel applications.

It was written for a competition that forced competitors to use VBA without
additional add-ins or COM objects.

Example
=======
py2vba can currently handle a variety of Python to VBA translation. The
following Python code is converted to a collection of VBA modules::

    from py2vba.convert import vbmeta
    from py2vba.vbast import String, Integer, Collection

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

PyMain.bas
----------
::

    Public Function test() As String
        Dim c As Company
        Set c = Company_ctor_(NewCollection(Person_ctor_("James", 27), Person_ctor_("Bob", 22)))
        test = c.employees(1 + 1).name
    End Function

    Private Function NewCollection(ParamArray params() As Variant) As Collection
        Dim p As Variant
        
        Set NewCollection = New Collection
        For Each p In params
            NewCollection.Add p
        Next p
    End Function

    Private Function NewDictionary(ParamArray params() As Variant) As Dictionary
        Dim k, v As Variant
        Dim i As Integer
        
        Debug.Assert (UBound(params) + 1) Mod 2 = 0
        Set NewDictionary = New Dictionary
        For i = LBound(params) To UBound(params) Step 2
            NewDictionary.Add params(i), params(i + 1)
        Next i
    End Function

PyMaincls_support.bas
---------------------
::

    Public Function Person_ctor_(name As String, age As Integer) As Person
        Set Person_ctor_ = New Person
        Person_ctor_.init__ name, age
    End Function
    Public Function Company_ctor_(employees As Collection) As Company
        Set Company_ctor_ = New Company
        Company_ctor_.init__ employees
    End Function

Person.cls
----------
::

    Public age as Integer
    Public name as String
    Public Function init__(name As String, age As Integer) As Variant
        Me.name = name
        Me.age = age
    End Function

Company.cls
-----------
::

    Public employees as Collection
    Public Function init__(employees As Collection) As Variant
        Set Me.employees = employees
    End Function
