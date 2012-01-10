class NodeWalkerError(Exception):
    pass

def visitor(node):
    def visitor_decorator(fcn):
        fcn.handles_node = node
        return fcn
    return visitor_decorator

class NodeWalker(object):
    def __init__(self):
        self._visitor_map = dict((x.handles_node, x) for x in self.__class__.__dict__.values() if hasattr(x, 'handles_node'))

    def walk(self, node):
        for cls in node.__class__.__mro__:
            if cls in self._visitor_map:
                return self._visitor_map[cls](self, node)
        raise NodeWalkerError('Cannot find walker handler associated with %r.' % (node,))
