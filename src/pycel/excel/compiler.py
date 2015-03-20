
class Context(object):
    """A small context object that nodes in the AST can use to emit code"""
    def __init__(self,curcell,excel):
        # the current cell for which we are generating code
        self.curcell = curcell
        # a handle to an excel instance
        self.excel = excel


class ExcelCompiler(object):
    """Class responsible for taking an Excel spreadsheet and compiling it to a Spreadsheet instance
       that can be serialized to disk, and executed independently of excel.
       
       Must be run on Windows as it requires a COM link to an Excel instance.
       """
       
    def __init__(self, filename=None, excel=None, *args,**kwargs):

        super(ExcelCompiler,self).__init__()
        self.filename = filename
        
        if excel:
            # if we are running as an excel addin, this gets passed to us
            self.excel = excel
        else:
            # TODO: use a proper interface so we can (eventually) support loading from file (much faster)  Still need to find a good lib though.
            self.excel = ExcelComWrapper(filename=filename)
            self.excel.connect()
            
        self.log = logging.getLogger("decode.{0}".format(self.__class__.__name__))
        
    def cell2code(self,cell):
        """Generate python code for the given cell"""
        if cell.formula:
            e = shunting_yard(cell.formula or str(cell.value))
            ast,root = build_ast(e)
            code = root.emit(ast,context=Context(cell,self.excel))
        else:
            ast = None
            code = str('"' + cell.value + '"' if isinstance(cell.value,unicode) else cell.value)
            
        return code,ast

    def add_node_to_graph(self,G, n):
        G.add_node(n)
        G.node[n]['sheet'] = n.sheet
        
        if isinstance(n,Cell):
            G.node[n]['label'] = n.col + str(n.row)
        else:
            #strip the sheet
            G.node[n]['label'] = n.address()[n.address().find('!')+1:]
            
    def gen_graph(self, seed, sheet=None):
        """Given a starting point (e.g., A6, or A3:B7) on a particular sheet, generate
           a Spreadsheet instance that captures the logic and control flow of the equations."""
        
        # starting points
        cursheet = sheet if sheet else self.excel.get_active_sheet()
        self.excel.set_sheet(cursheet)
        
        seeds,nr,nc = Cell.make_cells(self.excel, seed, sheet=cursheet)
        seeds = list(flatten(seeds))
        
        print "Seed %s expanded into %s cells" % (seed,len(seeds))
        
        # only keep seeds with formulas or numbers
        seeds = [s for s in seeds if s.formula or isinstance(s.value,(int,float))]

        print "%s filtered seeds " % len(seeds)
        
        # cells to analyze: only formulas
        todo = [s for s in seeds if s.formula]

        print "%s cells on the todo list" % len(todo)

        # map of all cells
        cellmap = dict([(x.address(),x) for x in seeds])
        
        # directed graph
        G = nx.DiGraph()

        # match the info in cellmap
        for c in cellmap.itervalues(): self.add_node_to_graph(G, c)

        while todo:
            c1 = todo.pop()
            
            print "Handling ", c1.address()
            
            # set the current sheet so relative addresses resolve properly
            if c1.sheet != cursheet:
                cursheet = c1.sheet
                self.excel.set_sheet(cursheet)
            
            # parse the formula into code
            pystr,ast = self.cell2code(c1)

            # set the code & compile it (will flag problems sooner rather than later)
            c1.python_expression = pystr
            c1.compile()                
            
            # get all the cells/ranges this formula refers to
            deps = [x.tvalue.replace('$','') for x in ast.nodes() if isinstance(x,RangeNode)]
            
            # remove dupes
            deps = uniqueify(deps)
            
            for dep in deps:
                
                # if the dependency is a multi-cell range, create a range object
                if is_range(dep):
                    # this will make sure we always have an absolute address
                    rng = CellRange(dep,sheet=cursheet)
                    
                    if rng.address() in cellmap:
                        # already dealt with this range
                        # add an edge from the range to the parent
                        G.add_edge(cellmap[rng.address()],cellmap[c1.address()])
                        continue
                    else:
                        # turn into cell objects
                        cells,nrows,ncols = Cell.make_cells(self.excel,dep,sheet=cursheet)

                        # get the values so we can set the range value
                        if nrows == 1 or ncols == 1:
                            rng.value = [c.value for c in cells]
                        else:
                            rng.value = [ [c.value for c in cells[i]] for i in range(len(cells)) ] 

                        # save the range
                        cellmap[rng.address()] = rng
                        # add an edge from the range to the parent
                        self.add_node_to_graph(G, rng)
                        G.add_edge(rng,cellmap[c1.address()])
                        # cells in the range should point to the range as their parent
                        target = rng
                else:
                    # not a range, create the cell object
                    cells = [Cell.resolve_cell(self.excel, dep, sheet=cursheet)]
                    target = cellmap[c1.address()]

                # process each cell                    
                for c2 in flatten(cells):
                    # if we havent treated this cell allready
                    if c2.address() not in cellmap:
                        if c2.formula:
                            # cell with a formula, needs to be added to the todo list
                            todo.append(c2)
                            #print "appended ", c2.address()
                        else:
                            # constant cell, no need for further processing, just remember to set the code
                            pystr,ast = self.cell2code(c2)
                            c2.python_expression = pystr
                            c2.compile()     
                            #print "skipped ", c2.address()
                        
                        # save in the cellmap
                        cellmap[c2.address()] = c2
                        # add to the graph
                        self.add_node_to_graph(G, c2)
                        
                    # add an edge from the cell to the parent (range or cell)
                    G.add_edge(cellmap[c2.address()],target)
            
        print "Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))

        sp = Spreadsheet(G,cellmap)
        
        return sp
