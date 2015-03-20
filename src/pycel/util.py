import collections
import re
import string


def is_range(address):
    return address.find(':') > 0 or not re.match('^\$?[A-Z]+[\d\$]+[A-Z\$]+[\d\$]+$', address)

def split_range(rng):
    if rng.find('!') > 0:
        sh,r = rng.split("!")
        start,end = r.split(':')
    else:
        sh = None
        start,end = rng.split(':')
    
    return sh,start,end
        
def split_address(address):
    sheet = None
    if address.find('!') > 0:
        sheet,address = address.split('!')
    
    #ignore case
    address = address.upper()
    
    # regular <col><row> format    
    if re.match('^[A-Z\$]+[\d\$]+$', address):
        col,row = filter(None,re.split('([A-Z\$]+)',address))
    # R<row>C<col> format
    elif re.match('^R\d+C\d+$', address):
        row,col = address.split('C')
        row = row[1:]
    # R[<row>]C[<col>] format
    elif re.match('^R\[\d+\]C\[\d+\]$', address):
        row,col = address.split('C')
        row = row[2:-1]
        col = col[2:-1]
    else:
        raise Exception('Invalid address format ' + address)
    
    return (sheet,col,row)

def resolve_range(rng, flatten=False, sheet=''):
    
    sh, start, end = split_range(rng)
    
    if sh and sheet:
        if sh != sheet:
            raise Exception("Mismatched sheets %s and %s" % (sh,sheet))
        else:
            sheet += '!'
    elif sh and not sheet:
        sheet = sh + "!"
    elif sheet and not sh:
        sheet += "!"
    else:
        pass

    # single cell, no range
    if not is_range(rng):  return ([sheet + rng],1,1)

    sh, start_col, start_row = split_address(start)
    sh, end_col, end_row = split_address(end)
    start_col_idx = col2num(start_col)
    end_col_idx = col2num(end_col);

    start_row = int(start_row)
    end_row = int(end_row)

    # single column
    if  start_col == end_col:
        nrows = end_row - start_row + 1
        data = [ "%s%s%s" % (s,c,r) for (s,c,r) in zip([sheet]*nrows,[start_col]*nrows,range(start_row,end_row+1))]
        return data,len(data),1
    
    # single row
    elif start_row == end_row:
        ncols = end_col_idx - start_col_idx + 1
        data = [ "%s%s%s" % (s,num2col(c),r) for (s,c,r) in zip([sheet]*ncols,range(start_col_idx,end_col_idx+1),[start_row]*ncols)]
        return data,1,len(data)
    
    # rectangular range
    else:
        cells = []
        for r in range(start_row,end_row+1):
            row = []
            for c in range(start_col_idx,end_col_idx+1):
                row.append(sheet + num2col(c) + str(r))
                
            cells.append(row)
    
        if flatten:
            # flatten into one list
            l = flatten(cells)
            return l,1,len(l)
        else:
            return cells, len(cells), len(cells[0]) 

# e.g., convert BA -> 53
def col2num(col):
    
    if not col:
        raise Exception("Column may not be empty")
    
    tot = 0
    for i,c in enumerate([c for c in col[::-1] if c != "$"]):
        if c == '$': continue
        tot += (ord(c)-64) * 26 ** i
    return tot

# convert back
def num2col(num):
    
    if num < 1:
        raise Exception("Number must be larger than 0: %s" % num)
    
    s = ''
    q = num
    while q > 0:
        (q,r) = divmod(q,26)
        if r == 0:
            q = q - 1
            r = 26
        s = string.ascii_uppercase[r-1] + s
    return s

def address2index(a):
    sh,c,r = split_address(a)
    return (col2num(c),int(r))

def index2addres(c,r,sheet=None):
    return "%s%s%s" % (sheet + "!" if sheet else "", num2col(c), r)



def flatten(l):
    for el in l:
        if isinstance(el, collections.Iterable) and not isinstance(el, basestring):
            for sub in flatten(el):
                yield sub
        else:
            yield el

def uniqueify(seq):
    seen = set()
    seen_add = seen.add
    return [ x for x in seq if x not in seen and not seen_add(x)]

