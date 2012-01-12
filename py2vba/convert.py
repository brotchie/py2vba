from nodewalker import NodeWalker, visitor, NodeWalkerError
import _ast, ast
import vbast

class PythonASTWalkerError(NodeWalkerError):
    pass

BINOP_MAP = {
    _ast.Add : '+',
    _ast.Sub : '-',
}

VBMETA = 'vbmeta'

def vbmeta(**kwargs):
    def vbmeta_decorator(fcn):
        fcn.vbmeta = dict(**kwargs)
        return fcn
    return vbmeta_decorator

def _extract_vbmeta_details(call):
    return dict((kw.arg, kw.value.id) for kw in
                    call.keywords)

class PythonASTWalker(NodeWalker):
    def __init__(self):
        super(PythonASTWalker, self).__init__()

        # State
        self._in_functiondef = None
        self._in_vbfunction = None
        self._in_vbmodule = None
        
        # Types
        self._types = {}
        for type in vbast.BUILTIN_TYPES:
            self.register_type(type)

    def register_type(self, typeobj):
        self._types[typeobj.name] = typeobj

    @visitor(_ast.Module)
    def visit_module(self, module):
        vbmodule = vbast.ProceduralModule('PyMain')
        self._in_vbmodule = vbmodule
        for c in module.body:
            if isinstance(c, _ast.FunctionDef):
                vbmodule.code.append(self.visit_functiondef(c))
            else:
                raise PythonASTWalkerError('Unrecognized Python AST node: %r', c)
        self._in_vbmodule = None
        return vbmodule

    @visitor(_ast.FunctionDef)
    def visit_functiondef(self, functiondef):
        if self._in_functiondef:
            raise PythonASTWalkerError('Cannot handle nested functiondefs at the moment,')

        # Check if vbmeta in decorators.
        vbmeta_decorators = [d for d in functiondef.decorator_list if
                                isinstance(d, _ast.Call) and
                                d.func.id == VBMETA]

        typeinfo = {}
        for d in vbmeta_decorators:
            typeinfo.update(_extract_vbmeta_details(d))

        self._in_functiondef = functiondef

        rettype = self._types.get(typeinfo.get('rettype', None), vbast.Variant)
        
        vbfunction = vbast.Function(functiondef.name,
                                    self._build_args(functiondef, typeinfo),
                                    rettype, [])
        self._in_vbmodule.function_namespace[vbfunction.name] = vbfunction
        assert self._in_vbfunction is None
        self._in_vbfunction = vbfunction

        vbfunction.statements.extend(sum([self.walk(c) for c in functiondef.body], []))

        # Check locals.
        vbfunction.statements = [vbast.DimStatement(name, type) for name, type in
                                    vbfunction.locals.iteritems()] + vbfunction.statements

        self._in_vbfunction = None
        self._in_functiondef = None
        return vbfunction

    @visitor(_ast.Assign)
    def visit_assign(self, assign):
        if len(assign.targets) > 1:
            raise PythonASTWalkerError('Cannot handle more than 1 assignment target.')

        lexpression = self.walk(assign.targets[0])
        if isinstance(assign.value, _ast.List):
            # Handle List literal.
            assert lexpression.name not in self._in_vbfunction.locals
            self._in_vbfunction.locals[lexpression.name] = vbast.Collection
            statements = [vbast.SetStatement(lexpression, vbast.NewExpression(vbast.Collection))]
            for element in assign.value.elts:
                statements.append(vbast.CallStatement(
                    vbast.MemberAccessExpression(lexpression, vbast.SimpleNameExpression('Add')),
                    [self.walk(element)]))
            return statements
        elif isinstance(assign.value, _ast.Dict):
            # Handle Dictionary literal.
            assert lexpression.name not in self._in_vbfunction.locals
            self._in_vbfunction.locals[lexpression.name] = vbast.Dictionary
            statements = [vbast.SetStatement(lexpression, vbast.NewExpression(vbast.Dictionary))]
            for (k,v) in zip(assign.value.keys, assign.value.values):
                statements.append(vbast.LetStatement(
                        vbast.IndexExpression(lexpression, [self.walk(k)]),
                        self.walk(v)))
            return statements
        else:
            raise PythonASTWalkerError('Unsupported RHS of assignment %r.' % (assign.value,))

    @visitor(_ast.Str)
    def visit_str(self, str):
        return vbast.StringLiteral(str.s)

    @visitor(_ast.Num)
    def visit_num(self, num):
        return vbast.IntegerLiteral(num.n)

    @visitor(_ast.Subscript)
    def visit_subscript(self, ss):
        lexpression = self.walk(ss.value)
        assert self._in_vbfunction.locals.get(lexpression.name) in [vbast.Collection, vbast.Dictionary]
        if isinstance(ss.slice.value, _ast.Num):
            return vbast.IndexExpression(lexpression,
                    [vbast.OperatorExpression('+', vbast.IntegerLiteral(ss.slice.value.n), 
                                             vbast.IntegerLiteral(1))])
        elif isinstance(ss.slice.value, _ast.Str):
            return vbast.IndexExpression(lexpression, [vbast.StringLiteral(ss.slice.value.s)])
        else:
            raise PythonASTWalkerError('Can only handle Integer and String array indexing.')

    @visitor(_ast.Return)
    def visit_return(self, ret):
        assert self._in_vbfunction
        return [vbast.LetStatement(
                 vbast.SimpleNameExpression(self._in_vbfunction.name),
                 self.walk(ret.value))]

    @visitor(_ast.BinOp)
    def visit_binop(self, binop):
        if binop.op.__class__ in BINOP_MAP:
            return vbast.OperatorExpression(BINOP_MAP[binop.op.__class__],
                    self.walk(binop.left), 
                    self.walk(binop.right))
        else:
            raise PythonASTWalkerError('Unhandled binary operation %s.' % (binop.op,))

    def _build_args(self, functiondef, typeinfo):
        args = []

        def process_arg(arg):
            argtype = self._types.get(typeinfo.get(arg.name), vbast.Variant)
            return vbast.Parameter(arg, argtype)

        return [process_arg(self.walk(a)) for a in functiondef.args.args]

    @visitor(_ast.Name)
    def visit_name(self, name):
        return vbast.SimpleNameExpression(name.id)

def build_ast_from_code(code):
    return compile(code, '<unknown>', 'exec', ast.PyCF_ONLY_AST)

def build_module_from_code(code):
    return compile(code, '<unknown>', 'exec')


if __name__ == '__main__':
    CODE = '''
@vbmeta(x=Integer, y=Integer, rettype=Integer)
def add(x, y):
    return x + y

def sub(x,y):
    return x - y

'''

    codeast = build_ast_from_code(CODE)
    codemodule = build_module_from_code(CODE)
    walker = PythonASTWalker()
    walker.register_type(vbast.Integer)
    result = walker.walk(codeast)
    print '\n'.join(result.as_code())
