"""
Nodes to represent a VBA program as an AST.

"""

PUBLIC = 'Public'
PRIVATE = 'Private'
FRIEND = 'Friend'
GLOBAL = 'Global'

STATIC = 'Static'

def indent(items):
    return ['\t' + item for item in items]

class VBType(object):
    _is_object_type = False

    @classmethod
    def is_object_type(cls):
        return cls._is_object_type

class ValueType(VBType):
    _is_object_type = False

class ObjectType(VBType):
    _is_object_type = True

class NamedValueType(ValueType):
    def __init__(self, name):
        self.name = name

class NamedObjectType(ObjectType):
    def __init__(self, name):
        self.name = name

Dictionary = NamedObjectType('Dictionary')
Collection = NamedObjectType('Collection')
Object = NamedObjectType('Object')
Integer = NamedValueType('Integer')

class VariantType(VBType):
    name = 'Variant'

Variant = VariantType()

BUILTIN_TYPES = [
    Dictionary, Object, Integer, Variant,
    Collection
]

class ASTNode(object):
    def as_code(self):
        """
        Returns the VBA code for this node and child nodes
        as a list of lines.

        """
        raise NotImplementedError()

    def _reduce_as_code(self, nodes):
        return sum([node.as_code() for node in nodes], [])

class Module(ASTNode):
    """
    Top level module.

    """
    @property
    def attributes(self):
        raise NotImplementedError()

    def _gen_module_header(self):
        return ['Attribute %s = "%s"' % (name, value) for
                    (name, value) in self.attributes]

class ProceduralModule(Module):
    def __init__(self, name):
        self.name = name
        self.directives = []
        self.declarations = []
        self.code = []
        self.function_namespace = {}

    @property
    def attributes(self):
        return [('VB_Name', self.name)]

    def as_code(self):
        return (self._gen_module_header() + [''] +
                self._reduce_as_code(self.directives) + 
                self._reduce_as_code(self.declarations) +
                self._reduce_as_code(self.code))

class ClassModule(Module):
    pass

class ModuleDirective(ASTNode):
    pass

class OptionExplicitDirective(ModuleDirective):
    def as_code(self):
        return ['Option Explicit']

class Procedure(ASTNode):
    pass

class Subroutine(Procedure):
    def __init__(self, name, parameters, statements, scope=PUBLIC, static=False):
        self.name = name
        self.scope = scope
        self.static = static
        self.parameters = parameters
        self.statements = statements
        self.locals = {}

    def as_code(self):
        paramlist = ', '.join(self._reduce_as_code(self.parameters))
        scope = self.scope
        if self.static:
            scope += ' ' + STATIC
        return (['%s Sub %s(%s)' % (scope, self.name, paramlist)] +
                indent(self._reduce_as_code(self.statements)) +  
                ['End Sub'])

class Function(Procedure):
    def __init__(self, name, parameters, rettype, statements=None, scope=PUBLIC, static=False):
        self.name = name
        self.parameters = parameters
        self.rettype = rettype
        self.scope = scope
        self.static = static
        self.statements = statements or []
        self.locals = {}

    def as_code(self):
        paramlist = ', '.join(self._reduce_as_code(self.parameters))
        scope = self.scope
        if self.static:
            scope += ' ' + STATIC
        return (['%s Function %s(%s) As %s' % (scope, self.name, paramlist, self.rettype.name)] +
                indent(self._reduce_as_code(self.statements)) +  
                ['End Function'])

class Parameter(ASTNode):
    def __init__(self, name, vbtype=Variant):
        self.name = name
        self.vbtype = vbtype

    def as_code(self):
        return ['%s As %s' % (self.name.as_code(), self.vbtype.name)]

    def __repr__(self):
        return 'Parameter(%r, %r)' % (self.name, self.vbtype)

class Statement(ASTNode):
    pass

class CallStatement(Statement):
    def __init__(self, lexpression, parameters):
        self.lexpression = lexpression
        self.parameters = parameters

    def as_code(self):
        return ['%s %s' % (self.lexpression.as_code(), ', '.join(p.as_code() for p in self.parameters))]

class IfStatement(Statement):
    def __init__(self, expression, thenblock, elseifblocks, elseblock):
        self.expression = expression
        self.thenblock = thenblock
        self.elseifblocks = elseifblocks
        self.elseblock = elseblock

class DimStatement(Statement):
    def __init__(self, name, vbtype, static=False):
        self.name = name
        self.vbtype = vbtype
        self.static = static

    def as_code(self):
        return ['Dim %s As %s' % (self.name, self.vbtype.name)]

class LetStatement(Statement):
    def __init__(self, lexpression, expression):
        self.lexpression = lexpression
        self.expression = expression

    def as_code(self):
        return ['%s = %s' % (''.join(self.lexpression.as_code()),
                             ''.join(self.expression.as_code()))]

class SetStatement(Statement):
    def __init__(self, lexpression, expression):
        self.lexpression = lexpression
        self.expression = expression

    def as_code(self):
        return ['Set %s = %s' % (''.join(self.lexpression.as_code()),
                                 ''.join(self.expression.as_code()))]

class Expression(ASTNode):
    pass

class OperatorExpression(Expression):
    def __init__(self, binop, left, right):
        self.binop = binop
        self.left = left
        self.right = right

    def as_code(self):
        return '%s %s %s' % (self.left.as_code(),
                             self.binop,
                             self.right.as_code())
class IndexExpression(Expression):
    def __init__(self, lexpression, args):
        self.lexpression = lexpression
        self.args = args

    def as_code(self):
        return '%s(%s)' % (self.lexpression.as_code(),
                           ','.join(a.as_code() for a in self.args))

class SimpleNameExpression(Expression):
    def __init__(self, name):
        self.name = name

    def as_code(self):
        return self.name

class ValueExpression(Expression):
    def __init__(self, code):
        self.code = code

    def as_code(self):
        return self.code

class LExpression(Expression):
    def __init__(self, name):
        self.name = name

    def as_code(self):
        return [self.name]

class NewExpression(Expression):
    def __init__(self, vbtype):
        self.vbtype = vbtype

    def as_code(self):
        return 'New %s' % (self.vbtype.name,)

class MemberAccessExpression(Expression):
    def __init__(self, lexpression, right):
        self.lexpression = lexpression
        self.right = right

    def as_code(self):
        return '%s.%s' % (self.lexpression.as_code(), self.right.as_code())

class StringLiteral(ASTNode):
    def __init__(self, value):
        self.value = value

    def as_code(self):
        return '"%s"' % (self.value,)

class IntegerLiteral(ASTNode):
    def __init__(self, value):
        self.value = value

    def as_code(self):
        return '%d' % (self.value,)


if __name__ == '__main__':
    m = ProceduralModule('Main')
    m.directives = [OptionExplicitDirective()]
    Integer = NamedVBType('Integer')
    statements = [DimStatement('x', NamedVBType('Dictionary'))]
    m.code = [Function('TestSub', [Parameter('x', Integer), Parameter('y', Integer)], Integer, statements, scope=PRIVATE, static=False)]
    print '\n'.join(m.as_code())
