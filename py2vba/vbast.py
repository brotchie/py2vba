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
    pass

class Variant(VBType):
    def __init__(self):
        self.name = 'Variant'

class NamedVBType(VBType):
    def __init__(self, name):
        self.name = name

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

    def as_code(self):
        paramlist = ', '.join(self._reduce_as_code(self.parameters))
        scope = self.scope
        if self.static:
            scope += ' ' + STATIC
        return ['%s Sub %s(%s)' % (scope, self.name, paramlist),
                'End Sub']

class Function(Procedure):
    def __init__(self, name, parameters, rettype, statements=None, scope=PUBLIC, static=False):
        self.name = name
        self.parameters = parameters
        self.rettype = rettype
        self.scope = scope
        self.static = static
        self.statements = statements or []

    def as_code(self):
        paramlist = ', '.join(self._reduce_as_code(self.parameters))
        scope = self.scope
        if self.static:
            scope += ' ' + STATIC
        return (['%s Function %s(%s) As %s' % (scope, self.name, paramlist, self.rettype.name)] +
                indent(self._reduce_as_code(self.statements)) +  
                ['End Function'])

class Parameter(ASTNode):
    def __init__(self, name, vbtype=Variant()):
        self.name = name
        self.vbtype = vbtype

    def as_code(self):
        return ['%s As %s' % (self.name.as_code(), self.vbtype.name)]

class Statement(ASTNode):
    pass

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
        return ['%s %s %s' % (self.left.as_code(),
                             self.binop,
                             self.right.as_code())]

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

if __name__ == '__main__':
    m = ProceduralModule('Main')
    m.directives = [OptionExplicitDirective()]
    Integer = NamedVBType('Integer')
    statements = [DimStatement('x', NamedVBType('Dictionary'))]
    m.code = [Function('TestSub', [Parameter('x', Integer), Parameter('y', Integer)], Integer, statements, scope=PRIVATE, static=False)]
    print '\n'.join(m.as_code())
