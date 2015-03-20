

from networkx.classes.digraph import DiGraph
from logging import info, debug

from .functions import (
    FUNCTION_MAP,
)

from .util import (
    is_range,
    split_address,
    split_range,
)

class Operator:
    """Small wrapper class to manage operators during shunting yard"""
    def __init__(self,value,precedence,associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity


class ASTNode(object):
    """A generic node in the AST"""
    
    def __init__(self,token):
        super(ASTNode,self).__init__()
        self.token = token
    def __str__(self):
        return self.token.tvalue
    def __getattr__(self,name):
        return getattr(self.token,name)

    def children(self,ast):
        args = ast.predecessors(self)
        args = sorted(args,key=lambda x: ast.node[x]['pos'])
        #args.reverse()
        return args

    def parent(self,ast):
        args = ast.successors(self)
        return args[0] if args else None
    
    def emit(self,ast,context=None):
        """Emit code"""
        self.token.tvalue
    
class OperatorNode(ASTNode):
    def __init__(self,*args):
        super(OperatorNode,self).__init__(*args)
        
        # convert the operator to python equivalents
        self.opmap = {
                 "^":"**",
                 "=":"==",
                 "&":"+",
                 "":"+" #union
                 }

    def emit(self,ast,context=None):
        xop = self.tvalue
        
        # Get the arguments
        args = self.children(ast)
        
        op = self.opmap.get(xop,xop)
        
        if self.ttype == "operator-prefix":
            return "-" + args[0].emit(ast,context=context)

        parent = self.parent(ast)
        # dont render the ^{1,2,..} part in a linest formula
        #TODO: bit of a hack
        if op == "**":
            if parent and parent.tvalue.lower() == "linest": 
                return args[0].emit(ast,context=context)

        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        if op == "<" or op == "<=":
            aa = args[0].emit(ast,context=context)
            ss = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + args[1].emit(ast,context=context)
        elif op == ">" or op == ">=":
            aa = args[1].emit(ast,context=context)
            ss =  args[0].emit(ast,context=context) + op + "(" + aa + " if " + aa + " is not None else float('inf'))"
        else:
            ss = args[0].emit(ast,context=context) + op + args[1].emit(ast,context=context)
        
        #avoid needless parentheses
        if parent and not isinstance(parent,FunctionNode):
            ss = "("+ ss + ")" 
        
        return ss

class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)
    def emit(self,ast,context=None):
        t = self.tsubtype
        
        if t == "logical":
            return str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            #if the string contains quotes, escape them
            val = self.tvalue.replace('"','\\"')
            return '"' + val + '"'
        else:
            return str(self.tvalue)

class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""
    def __init__(self,*args):
        super(RangeNode,self).__init__(*args)
    
    def emit(self,ast,context=None):
        # resolve the range into cells
        rng = self.tvalue.replace('$','')
        return rng

class FunctionNode(ASTNode):
    """AST node representing a function call"""
    def __init__(self,*args):
        super(FunctionNode,self).__init__(*args)
        self.numargs = 0

        # map  excel functions onto their python equivalents
        self.funmap = FUNCTION_MAP
        
    def emit(self,ast,context=None):
        fun = self.tvalue.lower()
        str = ''

        # Get the arguments
        args = self.children(ast)
        
        if fun == "atan2":
            # swap arguments
            str = "atan2(%s,%s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context))
        elif fun == "pi":
            # constant, no parens
            str = "pi"
        elif fun == "if":
            # inline the if
            if len(args) == 2:
                str = "%s if %s else 0" %(args[1].emit(ast,context=context),args[0].emit(ast,context=context))
            elif len(args) == 3:
                str = "(%s if %s else %s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context),args[2].emit(ast,context=context))
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":
            str += '['
            if len(args) == 1:
                # only one row
                str += args[0].emit(ast,context=context)
            else:
                # multiple rows
                str += ",".join(['[' + n.emit(ast,context=context) + ']' for n in args])
                     
            str += ']'
        elif fun == "arrayrow":
            #simply create a list
            str += ",".join([n.emit(ast,context=context) for n in args])
        elif fun == "linest" or fun == "linestmario":
            
            str = fun + "(" + ",".join([n.emit(ast,context=context) for n in args])

            if not context:
                degree,coef = -1,-1
            else:
                #linests are often used as part of an array formula spanning multiple cells,
                #one cell for each coefficient.  We have to figure out where we currently are
                #in that range
                degree,coef = get_linest_degree(context.excel,context.curcell)
                
            # if we are the only linest (degree is one) and linest is nested -> return vector
            # else return the coef.
            if degree == 1 and self.parent(ast):
                if fun == "linest":
                    str += ",degree=%s)" % degree
                else:
                    str += ")"
            else:
                if fun == "linest":
                    str += ",degree=%s)[%s]" % (degree,coef-1)
                else:
                    str += ")[%s]" % (coef-1)

        elif fun == "and":
            str = "all([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        elif fun == "or":
            str = "any([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        else:
            # map to the correct name
            f = self.funmap.get(fun,fun)
            str = f + "(" + ",".join([n.emit(ast,context=context) for n in args]) + ")"

        return str


def create_node(t):
    """Simple factory function"""
    if t.ttype == "operand":
        if t.tsubtype == "range":
            return RangeNode(t)
        else:
            return OperandNode(t)
    elif t.ttype == "function":
        return FunctionNode(t)
    elif t.ttype.startswith("operator"):
        return OperatorNode(t)
    else:
        return ASTNode(t)

def build_ast(expression):
    """build an AST from an Excel formula expression in reverse polish notation"""
    
    #use a directed graph to store the tree
    G = DiGraph()
    
    stack = []
    
    for n in expression:
        # Since the graph does not maintain the order of adding nodes/edges
        # add an extra attribute 'pos' so we can always sort to the correct order
        if isinstance(n,OperatorNode):
            if n.ttype == "operator-infix":
                arg2 = stack.pop()
                arg1 = stack.pop()
                G.add_node(arg1,{'pos':1})
                G.add_node(arg2,{'pos':2})
                G.add_edge(arg1, n)
                G.add_edge(arg2, n)
            else:
                arg1 = stack.pop()
                G.add_node(arg1,{'pos':1})
                G.add_edge(arg1, n)
                
        elif isinstance(n,FunctionNode):
            args = [stack.pop() for _ in range(n.num_args)]
            args.reverse()
            for i,a in enumerate(args):
                G.add_node(a,{'pos':i})
                G.add_edge(a,n)
            #for i in range(n.num_args):
            #    G.add_edge(stack.pop(),n)
        else:
            G.add_node(n,{'pos':0})

        stack.append(n)
        
    return G,stack.pop()


