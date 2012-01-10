from nodewalker import NodeWalker, visitor, NodeWalkerError
import _ast, ast
import vbast

class PythonASTWalkerError(NodeWalkerError):
    pass

BINOP_MAP = {
    _ast.Add : '+',
    _ast.Sub : '-',
}

class PythonASTWalker(NodeWalker):
    def __init__(self):
        super(PythonASTWalker, self).__init__()

        # State
        self._in_functiondef = None
        self._in_vbfunction = None

    @visitor(_ast.Module)
    def visit_module(self, module):
        vbmodule = vbast.ProceduralModule('PyMain')
        for c in module.body:
            if isinstance(c, _ast.FunctionDef):
                vbmodule.code.append(self.visit_functiondef(c))
            else:
                raise PythonASTWalkerError('Unrecognized Python AST node: %r', c)
        return vbmodule

    @visitor(_ast.FunctionDef)
    def visit_functiondef(self, functiondef):
        if self._in_functiondef:
            raise PythonASTWalkerError('Cannot handle nested functiondefs at the moment,')

        self._in_functiondef = functiondef
        
        vbfunction = vbast.Function(functiondef.name,
                                    self._build_args(functiondef),
                                    vbast.Variant(), [])
        assert self._in_vbfunction is None
        self._in_vbfunction = vbfunction

        vbfunction.statements.extend(self.walk(c) for c in functiondef.body)

        self._in_vbfunction = None
        self._in_functiondef = None
        return vbfunction

    @visitor(_ast.Return)
    def visit_return(self, ret):
        assert self._in_vbfunction
        return vbast.LetStatement(
                vbast.SimpleNameExpression(self._in_vbfunction.name),
                self.walk(ret.value))

    @visitor(_ast.BinOp)
    def visit_binop(self, binop):
        if binop.op.__class__ in BINOP_MAP:
            return vbast.OperatorExpression(BINOP_MAP[binop.op.__class__],
                    self.walk(binop.left), 
                    self.walk(binop.right))
        else:
            raise PythonASTWalkerError('Unhandled binary operation %s.' % (binop.op,))

    def _build_args(self, functiondef):
        args = []

        def process_arg(arg):
            return vbast.Parameter(arg, vbast.Variant())

        return [process_arg(self.walk(a)) for a in functiondef.args.args]

    @visitor(_ast.Name)
    def visit_name(self, name):
        return vbast.SimpleNameExpression(name.id)

def build_ast_from_code(code):
    return compile(code, '<unknown>', 'exec', ast.PyCF_ONLY_AST)

def build_module_from_code(code):
    return compile(code, '<unknown>', 'exec')

if __name__ == '__main__':
    CODE = '''def add(x, y):
    return x + y

def sub(x,y):
    return x - y

'''

    codeast = build_ast_from_code(CODE)
    codemodule = build_module_from_code(CODE)
    walker = PythonASTWalker()
    result = walker.walk(codeast)
    print '\n'.join(result.as_code())
