import cPickle

import networkx as nx

from networkx.drawing.nx_pydot import write_dot
from networkx.drawing.nx_pylab import draw, draw_circular
from networkx.readwrite.gexf import write_gexf


class Spreadsheet(object):
    def __init__(self,G,cellmap):
        super(Spreadsheet,self).__init__()
        self.G = G
        self.cellmap = cellmap
        self.params = None

    @staticmethod
    def load_from_file(fname):
        f = open(fname,'rb')
        obj = cPickle.load(f)
        #obj = load(f)
        return obj
    
    def save_to_file(self,fname):
        f = open(fname,'wb')
        cPickle.dump(self, f, protocol=2)
        f.close()

    def export_to_dot(self,fname):
        write_dot(self.G,fname)
                    
    def export_to_gexf(self,fname):
        write_gexf(self.G,fname)
    
    def plot_graph(self):
        import matplotlib.pyplot as plt

        pos=nx.spring_layout(self.G,iterations=2000)
        #pos=nx.spectral_layout(G)
        #pos = nx.random_layout(G)
        nx.draw_networkx_nodes(self.G, pos)
        nx.draw_networkx_edges(self.G, pos, arrows=True)
        nx.draw_networkx_labels(self.G, pos)
        plt.show()
    
    def set_value(self,cell,val,is_addr=True):
        if is_addr:
            cell = self.cellmap[cell]

        if cell.value != val:
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = val
        
    def reset(self, cell):
        if cell.value is None: return
        #print "resetting", cell.address()
        cell.value = None
        map(self.reset,self.G.successors_iter(cell)) 

    def print_value_tree(self,addr,indent):
        cell = self.cellmap[addr]
        print "%s %s = %s" % (" "*indent,addr,cell.value)
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def recalculate(self):
        for c in self.cellmap.values():
            if isinstance(c,CellRange):
                self.evaluate_range(c,is_addr=False)
            else:
                self.evaluate(c,is_addr=False)
                
    def evaluate_range(self,rng,is_addr=True):

        if is_addr:
            rng = self.cellmap[rng]

        # its important that [] gets treated ad false here
        if rng.value:
            return rng.value

        cells,nrows,ncols = rng.celladdrs,rng.nrows,rng.ncols

        if nrows == 1 or ncols == 1:
            data = [ self.evaluate(c) for c in cells ]
        else:
            data = [ [self.evaluate(c) for c in cells[i]] for i in range(len(cells)) ] 
        
        rng.value = data
        
        return data

    def evaluate(self,cell,is_addr=True):

        if is_addr:
            cell = self.cellmap[cell]
            
        # no formula, fixed value
        if not cell.formula or cell.value != None:
            #print "  returning constant or cached value for ", cell.address()
            return cell.value
        
        # recalculate formula
        # the compiled expression calls this function
        def eval_cell(address):
            return self.evaluate(address)
        
        def eval_range(rng):
            return self.evaluate_range(rng)
                
        try:
            #print "Evalling: %s, %s" % (cell.address(),cell.python_expression)
            vv = eval(cell.compiled_expression)
            #print "Cell %s evalled to %s" % (cell.address(),vv)
            if vv is None:
                print "WARNING %s is None" % (cell.address())
            cell.value = vv
        except Exception as e:
            if e.message.startswith("Problem evalling"):
                raise e
            else:
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression)) 
        
        return cell.value

