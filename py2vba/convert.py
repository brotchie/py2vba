from nodewalker import NodeWalker, visitor, NodeWalkerError
import _ast, ast
import vbast

class PythonASTWalkerError(NodeWalkerError):
    pass

BINOP_MAP = {
    _ast.Add : '+',
    _ast.Sub : '-',
    _ast.Mult : '*',
    _ast.Mod : 'Mod',
}

UNARYOP_MAP = {
    _ast.USub : '-',
}

COMPAREOP_MAP = {
    _ast.Gt : '>',
    _ast.Lt : '<',
    _ast.GtE : '>=',
    _ast.LtE : '<=',
    _ast.Eq : '=', 
}

BOOLOP_MAP = {
    _ast.And : 'And',
    _ast.Or : 'Or',
}

VBMETA = 'vbmeta'

def vbmeta(**kwargs):
    def vbmeta_decorator(fcn):
        fcn.vbmeta = dict(**kwargs)
        return fcn
    return vbmeta_decorator

def _extract_vbmeta_details(call):
    return [(kw.arg, kw.value.id) for kw in
                    call.keywords]

class PythonASTWalker(NodeWalker):
    def __init__(self):
        super(PythonASTWalker, self).__init__()

        # State
        self._in_vbfunction = None
        self._in_vbmodule = None
        self._assign_lexpression = None

        self._in_vbclassmodule = None
        self._selfname = None

        self._classnames = []
        
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
                vbmodule.code.append(self.walk(c))
            elif isinstance(c, _ast.ClassDef):
                self.walk(c)
            else:
                raise PythonASTWalkerError('Unrecognized Python AST node: %r', c)

        vbmodule.raw_code.append(vbast.COLLECTION_LITERAL_HELPERS)
        self._in_vbmodule = None
        return vbmodule

    def _extract_typeinfo_from_functiondef(self, functiondef):
        vbmeta_decorators = [d for d in functiondef.decorator_list if
                                isinstance(d, _ast.Call) and
                                d.func.id == VBMETA]

        rawtypeinfo = set()
        for d in vbmeta_decorators:
            rawtypeinfo.update(_extract_vbmeta_details(d))
        return {varname:self._types[typename] for varname, typename in rawtypeinfo}

    def _build_args(self, functiondef, typeinfo):
        if self._in_vbclassmodule:
            functiondefargs = functiondef.args.args[1:]
            selfname = functiondef.args.args[0].id
        else:
            functiondefargs = functiondef.args.args
            selfname = None

        def process_arg(arg):
            argtype = typeinfo.get(arg.name, vbast.Variant)
            return vbast.Parameter(arg, argtype)

        return [process_arg(self.walk(a)) for a in functiondefargs], selfname

    def _create_dim_statements(self, vardefs):
        return [vbast.DimDeclaration(name, type) for name,type in vardefs]

    def _create_and_add_class_ctor(self, functiondef, vbfunction, typeinfo):
        ctorname = self._in_vbclassmodule.name + '_ctor_'

        if not self._in_vbmodule.class_support_module:
            self._in_vbmodule.class_support_module = \
                    vbast.ProceduralModule(self._in_vbmodule.name + 'cls_support')

        self._in_vbmodule.class_support_module.code.append(
            vbast.Function(
                ctorname,
                vbfunction.parameters,
                self._in_vbclassmodule.vbtype(),
                [
                    vbast.SetStatement(
                        vbast.SimpleNameExpression(ctorname),
                        vbast.NewExpression(self._in_vbclassmodule.vbtype())
                    ),
                    vbast.CallStatement(
                        vbast.MemberAccessExpression(
                            vbast.SimpleNameExpression(ctorname),
                            vbast.SimpleNameExpression('init__'),
                        ),
                        [p.name for p in vbfunction.parameters]
                    )
                ]
            )
        )

        # Try to determine instance variables from assignments
        # within the __init__ body.
        instance_variables = {}
        for stmt in functiondef.body:
            if not isinstance(stmt, _ast.Assign):
                continue
            lhs = stmt.targets[0]
            isSimpleAssignment = isinstance(lhs, _ast.Attribute) and \
                    isinstance(lhs.value, _ast.Name)

            if isSimpleAssignment and lhs.value.id == self._selfname:
                if isinstance(stmt.value, _ast.Name):
                    rhsname = stmt.value.id
                else:
                    rhsname = None
                instance_variables[lhs.attr] = typeinfo.get(rhsname, vbast.Variant)

        # Create instance variable declarations.
        self._in_vbclassmodule.declarations += [vbast.PublicVariableDeclaration(name, type) 
                                                for name, type 
                                                in instance_variables.iteritems()]

    @visitor(_ast.FunctionDef)
    def visit_functiondef(self, functiondef):
        assert not self._in_vbfunction, 'Cannot handle nested functiondefs at the moment,'

        typeinfo = self._extract_typeinfo_from_functiondef(functiondef)
        rettype = typeinfo.get('rettype', vbast.Variant)
        args, self._selfname = self._build_args(functiondef, typeinfo)

        if functiondef.name == '__init__' and self._in_vbclassmodule:
            fname = 'init__'
        else:
            fname = functiondef.name
        
        vbfunction = vbast.Function(fname, args, rettype, [])

        if self._in_vbclassmodule:
            self._in_vbclassmodule.method_namespace[vbfunction.name] = vbfunction
        else:
            self._in_vbmodule.function_namespace[vbfunction.name] = vbfunction

        if fname == 'init__':
            self._create_and_add_class_ctor(functiondef, vbfunction, typeinfo)

        self._in_vbfunction = vbfunction

        body_statements = sum([self.walk(c) for c in functiondef.body], [])
        dim_statements = self._create_dim_statements(vbfunction.locals.iteritems())

        vbfunction.statements = dim_statements + body_statements

        self._in_vbfunction = None
        self._selfname = None

        return vbfunction

    @visitor(_ast.Assign)
    def visit_assign(self, assign):
        if len(assign.targets) > 1:
            raise PythonASTWalkerError('Cannot handle more than 1 assignment target.')

        lexpression = self.walk(assign.targets[0])
        rhs = self.walk(assign.value)
        if isinstance(lexpression, vbast.SimpleNameExpression) and \
                lexpression.name not in self._in_vbfunction.locals:
            self._in_vbfunction.locals[lexpression.name] = rhs.vbtype()

        if rhs.vbtype().is_object_type():
            assignment_statment = vbast.SetStatement
        else:
            assignment_statment = vbast.LetStatement

        return [assignment_statment(lexpression, rhs)]

    @visitor(_ast.Dict)
    def visit_dict(self, dict):
        return vbast.DictLiteral([(self.walk(k), self.walk(v)) for k,v in zip(dict.keys, dict.values)])

    @visitor(_ast.List)
    def visit_list(self, list):
        return vbast.ListLiteral([self.walk(e) for e in list.elts])

    @visitor(_ast.Str)
    def visit_str(self, str):
        return vbast.StringLiteral(str.s)

    @visitor(_ast.Num)
    def visit_num(self, num):
        return vbast.IntegerLiteral(num.n)

    @visitor(_ast.Subscript)
    def visit_subscript(self, ss):
        lexpression = self.walk(ss.value)
        if isinstance(ss.slice.value, _ast.Num):
            return vbast.IndexExpression(lexpression,
                    [vbast.BinOp('+', vbast.IntegerLiteral(ss.slice.value.n), 
                                             vbast.IntegerLiteral(1))])
        elif isinstance(ss.slice.value, _ast.Str):
            return vbast.IndexExpression(lexpression, [vbast.StringLiteral(ss.slice.value.s)])
        else:
            raise PythonASTWalkerError('Can only handle Integer and String array indexing.')

    @visitor(_ast.Return)
    def visit_return(self, ret):
        assert self._in_vbfunction
        if isinstance(self._in_vbfunction, vbast.Function):
            return_statement = vbast.ExitFunctionStatement()
        else:
            return_statement = vbast.ExitSubStatement()
        if self._in_vbfunction.rettype.is_object_type():
            assign_stmt_type = vbast.SetStatement
        else:
            assign_stmt_type = vbast.LetStatement
        return [assign_stmt_type(
                 vbast.SimpleNameExpression(self._in_vbfunction.name),
                 self.walk(ret.value)), return_statement]

    @visitor(_ast.BinOp)
    def visit_binop(self, binop):
        if binop.op.__class__ in BINOP_MAP:
            return vbast.BinOp(BINOP_MAP[binop.op.__class__],
                    self.walk(binop.left), 
                    self.walk(binop.right))
        else:
            raise PythonASTWalkerError('Unhandled binary operation %s.' % (binop.op,))

    @visitor(_ast.UnaryOp)
    def visit_unaryop(self, unaryop):
        return vbast.UnaryOp(UNARYOP_MAP[unaryop.op.__class__],
                self.walk(unaryop.operand))

    @visitor(_ast.Call)
    def visit_call(self, call):
        if call.func.id in self._classnames:
            expression = vbast.IndexExpression(
                    vbast.SimpleNameExpression(call.func.id + '_ctor_'),
                    [self.walk(a) for a in call.args])
            expression.set_vbtype(vbast.NamedObjectType(call.func.id))
        else:
            expression = vbast.IndexExpression(
                    self.walk(call.func),
                    [self.walk(a) for a in call.args])

        return expression

    @visitor(_ast.Attribute)
    def visit_attribute(self, attribute):
        return vbast.MemberAccessExpression(
                self.walk(attribute.value),
                vbast.SimpleNameExpression(attribute.attr))

    @visitor(_ast.ClassDef)
    def visit_classdef(self, classdef):
        self._in_vbclassmodule = vbast.ClassModule(classdef.name)
        self._classnames.append(classdef.name)
        self._in_vbmodule.support_modules.append(self._in_vbclassmodule)
        self._in_vbclassmodule.code.extend([self.walk(c) for c in classdef.body])
        self._in_vbclassmodule = None
    
    @visitor(_ast.Name)
    def visit_name(self, name):
        if name.id == self._selfname:
            expression = vbast.SimpleNameExpression('Me')
        else:
            expression = vbast.SimpleNameExpression(name.id)
            if self._in_vbfunction:
                if name.id in self._in_vbfunction.parameters_names:
                    expression.set_vbtype(self._in_vbfunction.get_parameter_type(name.id))

        return expression

    @visitor(_ast.If)
    def visit_ifstatement(self, ifstmt):
        return [vbast.IfStatement(
            self.walk(ifstmt.test),
            self._walk_block(ifstmt.body),
            [], self._walk_block(ifstmt.orelse))]

    @visitor(_ast.Compare)
    def visit_compare(self, compare):
        return vbast.BinOp(
                COMPAREOP_MAP[compare.ops[0].__class__],
                self.walk(compare.left),
                self.walk(compare.comparators[0]))

    @visitor(_ast.For)
    def visit_for(self, forstmt):
        # Currently only support interpreting range
        # with 2 arguments.
        assert forstmt.iter.func.id == 'range'
        assert len(forstmt.iter.args) == 2
        
        ifrom = forstmt.iter.args[0].n
        ito = forstmt.iter.args[1].n - 1

        if forstmt.target.id not in self._in_vbfunction.locals:
            self._in_vbfunction.locals[forstmt.target.id] = vbast.Integer

        return [vbast.ForStatement(
            self.walk(forstmt.target),
            self._walk_block(forstmt.body),
            vbast.IntegerLiteral(ifrom),
            vbast.IntegerLiteral(ito))]

    @visitor(_ast.AugAssign)
    def visit_augassign(self, augassign):
        return [vbast.LetStatement(
                self.walk(augassign.target),
                vbast.BinOp(BINOP_MAP[augassign.op.__class__],
                            self.walk(augassign.target),
                            self.walk(augassign.value)))]

    @visitor(_ast.BoolOp)
    def visit_boolop(self, boolop):
        expr = vbast.BinOp(
                BOOLOP_MAP[boolop.op.__class__],
                self.walk(boolop.values[0]),
                self.walk(boolop.values[1]))
        for i in range(2, len(boolop.values)):
            expr = vbast.BinOp(
                    BOOLOP_MAP[boolop.op.__class__],
                    expr, self.walk(boolop.values[i]))
        return expr

    @visitor(_ast.ListComp)
    def visit_listcomp(self, listcomp):
        assert len(listcomp.generators) == 1, 'List comp conversion only supports a single generator.'
        assert len(listcomp.generators[0].ifs) <= 1
        itervars = { self.walk(listcomp.generators[0].iter) }
        itervarnames = { n.name for n in itervars }
        target = self.walk(listcomp.generators[0].target)

        # Determine if there are any other variables to include as
        # as closure.
        othervars = { self.walk(n) for n in ast.walk(listcomp) 
                        if isinstance(n, _ast.Name) and
                           n.id not in itervarnames and
                           n.id != target.name }

        fname = '%s_listcomp_%i' % (self._in_vbfunction.name, len(self._in_vbfunction.listcomps))

        vbfunctionlocals = {p.name : p.vbtype for p in self._in_vbfunction.parameters}
        vbfunctionlocals.update(self._in_vbfunction.locals)

        parameters = [vbast.Parameter(n, vbfunctionlocals.get(n.name, vbast.Variant))
                        for n in othervars | itervars] 

        assert len(itervars) == 1
        itervar = list(itervars)[0]

        compexpr = self.walk(listcomp.elt)

        vbfunction = vbast.Function(fname, parameters, vbast.Collection, scope=vbast.PRIVATE)
        bodystmt = vbast.CallStatement(
                        vbast.MemberAccessExpression(
                            vbast.SimpleNameExpression(fname),
                            vbast.SimpleNameExpression('Add')),
                        [compexpr])

        if listcomp.generators[0].ifs:
            bodystmt = vbast.IfStatement(self.walk(listcomp.generators[0].ifs[0]),
                                         [bodystmt])

        vbfunction.statements += [
                                  vbast.DimDeclaration(target.name, vbast.Variant),
                                  vbast.SetStatement(vbast.SimpleNameExpression(fname),
                                                     vbast.NewExpression(vbast.Collection)),
                                  vbast.ForEachStatement(target, itervar, [bodystmt])]
        self._in_vbfunction.listcomps.append(vbfunction)

        expr =  vbast.IndexExpression(vbast.SimpleNameExpression(vbfunction.name),
                                     [param.name for param in parameters])
        expr.set_vbtype(vbast.Collection)
        return expr

    def _walk_block(self, block):
        return sum([self.walk(c) for c in block], [])

def build_ast_from_code(code):
    return compile(code, '<unknown>', 'exec', ast.PyCF_ONLY_AST)

def build_module_from_code(code):
    return compile(code, '<unknown>', 'exec')
