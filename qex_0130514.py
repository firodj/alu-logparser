import os,sys,re,shutil
import zipfile, codecs
# version: 14 mei 2013
# version: 26 jun 2014 -- normalize

# apply rules in given order!
rules_pattern = [
    #{ r'>\s+' : u'>'},                  # remove spaces after a tag opens or closes
    { r'\s+' : u' '},                   # replace consecutive spaces
    { r'\s*<br\s*/?>\s*' : u'\n'},      # newline after a <br>
    { r'</(div)\s*>\s*' : u'\n'},       # newline after </p> and </div> and <h1/>...
    { r'</(p|h\d)\s*>\s*' : u'\n\n'},   # newline after </p> and </div> and <h1/>...
    { r'<head>.*<\s*(/head|body)[^>]*>' : u'' },     # remove <head> to </head>
    { r'<a\s+href="([^"]+)"[^>]*>.*</a>' : r'\1' },  # show links instead of texts
    { r'[ \t]*<[^<]*?/?>' : u'' },            # remove remaining tags
    { r'^\s+' : u'' }                   # remove spaces at the beginning
]

rules_regex = list()
for rule in rules_pattern:
    for (k,v) in rule.items():
        regex = re.compile(k)
        rules_regex.append( {v: regex} )
        
def stripHTMLTags(html):
    """
    Strip HTML tags from any string and transfrom special entities
    http://www.codigomanso.com/en/2010/09/truco-manso-eliminar-tags-html-en-python/
    """
    global rules_regex
    text = html

    for rule in rules_regex:
        for (v,regex) in rule.items():
            text = regex.sub(v, text)

    # replace special strings
    special = {
        '&nbsp;' : ' ', '&amp;' : '&', '&quot;' : '"',
        '&lt;'   : '<', '&gt;'  : '>'
    }

    for (k,v) in special.items():
        text = text.replace (k, v)

    return text

def removetags(data):
    return re.sub(r'<[^<]*?>', '', data)

def safehtml(data):
    return data.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def extract_template_(template):
    st_xlsx = os.stat(template + '.xlsx')
    extract = False

    for f in ['/xl/sharedStrings.xml', '/xl/worksheets/sheet1.xml']:
        if os.path.exists('tmp/' + template + f):
            st_test = os.stat('tmp/' + template + f)
            if st_xlsx.st_mtime > st_test.st_mtime:
                extract = True
                break
        else:
            extract = True
            break

    if extract:
        print "DBG: Extracting template...", template
        z = zipfile.ZipFile(template + '.xlsx', 'r')
        z.extractall('tmp/' + template + '/')
    
def _makesure(s_path, fname, template):
    path = 'out/report/[' + template + ']_' + s_path
    
    full = path + fname
    subs = full.rsplit('/', 1)
    
    if not os.path.exists(subs[0]):
        os.makedirs(subs[0])
    if not os.path.exists(full):
        shutil.copyfile('tmp/' + template + fname, full)
    return full

def listall(path, exclude=[]):
    files = list()
    for f in os.listdir(path):
        r_f = path + '/' + f
        if os.path.isdir(r_f):
            files = files + listall(r_f, exclude)
        elif os.path.isfile(r_f):
            skip = False
            for e in exclude:
                if re.match('.*' + e + '$', r_f):
                    skip = True
                    print "DBG: Skipping...", r_f
                    break
            if not skip:
                files.append(r_f)
    return files

g_template_collections = dict()

class QuickTemplate:
    def __init__(self, template):
        #global g_template_collections

        self.name = template
        
        #if self.name not in g_template_collections:
        #    extract_template( self.name )
        #    self.files = listall( 'tmp/' + self.name, ['/xl/sharedStrings.xml', '/xl/worksheets/sheet1.xml'])
        #    g_template_collections[ self.name ] = self.files
        #else:
        #    self.files = g_template_collections[ self.name ]
            
    def extract(self, path):
        z = zipfile.ZipFile('templates/'+ self.name + '.xlsx', 'r')
        z.extractall(path)
        z.close()
        
class QuickCell:
    def __init__(self, cell_id, attrs, val):
        self.id = cell_id
        self.type = None
        self.style = None
        
        x = attrs.split('s=') # cell style
        attrs = x[0]
        if len(x) > 1:
            y = x[1].split(' ')
            self.style = y[0].strip('" ')
            if len(y) > 1: attrs = attrs + y[1]

        x = attrs.split('t=') # cell format
        attrs = x[0]
        if len(x) > 1:
            y = x[1].split(' ')
            self.type = y[0].strip('" ')
            if len(y) > 1: attrs = attrs + y[1]

        self.attrs = attrs.strip(' ')
        self.val = val

    def is_string(self):
        return self.val <> None and self.type == 's'
        
    def pret(self):
        attrs = ['r="%s"' % (self.id, ) ]
        if len(self.attrs) > 0: attrs.append( self.attrs )
        if self.style: attrs.append('s="%s"' % (self.style, ) )
        if self.type: attrs.append('t="%s"' % (self.type, ) )
        
        s = '<c ' + " ".join(attrs)
        
        if self.val:
            s = s + '><v>%s</v></c>' % (self.val,)
        else:
            s = s + '/>'
        return s    
    
class QuickRow:
    def __init__(self, row_id, attrs):
        self.id = row_id
        self.attrs = attrs.strip(' ')
        
        self.cells = dict()
        self.cells_key = list()

    def add_cell(self, cell_id, attrs, val):
        cell_obj = QuickCell(cell_id, attrs, val)

        self.cells[ cell_obj.id ] = cell_obj
        self.cells_key.append( cell_obj.id )

        return cell_obj

    def pret(self):
        attrs = [ 'r="%s"' % (self.id,) ]
        if len(self.attrs) > 0: attrs.append(self.attrs)
        s = '<row ' + ' '.join(attrs) + '>'
        self.cells_key.sort(cmp = cmp_coord)
        for cell_key in self.cells_key:
            cell_obj = self.cells[ cell_key ]
            s = s + cell_obj.pret()
        s = s + '</row>'
        return s

def cmp_coord(kx, ky):
    x = re.match("([A-Z]+)([0-9]+)", kx)
    y = re.match("([A-Z]+)([0-9]+)", ky)
    
    r = cmp( int(x.group(2)), int(y.group(2)) )
    if r <> 0: return r

    cl = cmp( len(x.group(1)), len(y.group(1)) )
    if cl <> 0: return cl
    
    c = cmp( x.group(1), y.group(1) )
    return c

class QuickString:
    def __init__(self, str_id, raw_xml=None, val_str=None, val_pre=False):
        self.id = str_id
        self.pre = None
        
        xml_attr = None
        if raw_xml == None:
            val_attr = ""
            if val_pre: val_attr = val_attr + ' xml:space="preserve"'
            val_str = safehtml(val_str)
            raw_xml = '<si><t%s>%s</t></si>' % (val_attr,val_str)
        self.raw_xml = raw_xml

        #self.xml = BeautifulSoup(raw_xml, "xml")
        #self.pure = self.xml.text
        #TESTING:ME
        self.pure = stripHTMLTags(raw_xml)
        
        m = re.search("<t\s+xml:space=(.+?)\s*>", self.raw_xml)
        if m: self.pre = m.group(1)
        
        self.ref = dict()

    def add_ref(self, cell_id):
        self.ref[cell_id] = True

    def del_ref(self, cell_id):
        del self.ref[cell_id]

    
class QuickExcel:
    def __init__(self, path, template):
        self.path = path
        self.templ = QuickTemplate(template)
        self.open_xlsx()
        
        #self.f_strings = makesure( self.path , '/xl/sharedStrings.xml', self.templ.name)
        
        f = codecs.open(self.f_strings, 'r', 'utf-8')
        data = f.read()
        f.close()

        m = re.search('<sst xmlns.*? count=(["0-9]+)', data)
        
        self.strings_count = int(m.group(1).strip('" '))
        self.strings = list()
        self.strings_index = dict()

        for m in re.finditer('<si>.*?</si>', data, re.S):
            qs = QuickString( len(self.strings), m.group(0))
            self.strings.append( qs )
            self.strings_index[qs.pure] = qs

        #self.f_sheet = makesure(self.path, '/xl/worksheets/sheet1.xml', self.templ.name)
        
        f = open(self.f_sheet, 'r')
        sheet = f.read()
        f.close()

        self.sheet_untouch = list()
        m = re.search('<sheetData>.+?</sheetData>', sheet)
        self.sheet_untouch.append( sheet[0:m.start()] )
        self.sheet_untouch.append( sheet[m.end():] )

        self.rows_key = list()
        self.rows = dict()

        self.dbg = m.group(0)
        
        i = 0
        for m in re.finditer('<row r=(["0-9]+)(.*?)>(.+?)</row>', m.group(0) ):
            row_obj = self.create_row( m.group(1).strip('" '), m.group(2) )
            for n in re.finditer('<c r=(["A-Z0-9]+)(.*?)(/>|><v>(.+?)</v></c>)', m.group(3)):
                cell_obj = row_obj.add_cell( n.group(1).strip('" '), n.group(2), n.group(4) )
                if cell_obj.is_string():
                    
                    qs = self.strings[ int(cell_obj.val) ]
                    qs.add_ref(cell_obj.id)
                    i = i + 1

        self.strings_count = i

    def create_row(self, row_id, attrs):
        row_obj = QuickRow( row_id, attrs )

        self.rows[row_obj.id] = row_obj
        self.rows_key.append(row_obj.id)
        return row_obj
                
    def save_strings(self):
        f = codecs.open(self.f_strings, 'w', 'utf-8')
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">' % (self.strings_count, len(self.strings)) )
        for qs in self.strings:
            f.write(qs.raw_xml)
        f.write('</sst>')
        f.close()

    def save_sheet(self):
        f = open(self.f_sheet, 'w')
        
        f.write( self.sheet_untouch[0] )
        f.write('<sheetData>')
        self.rows_key.sort(key = lambda x: int(x))
        for row_key in self.rows_key:
            row_obj = self.rows[row_key]
            f.write( row_obj.pret() )
        f.write('</sheetData>')
        f.write( self.sheet_untouch[1] )
        f.close()

    def get_cell(self, coordinate):
        m = re.match('([A-Z]+)([0-9]+)', coordinate)
        if not m: return False

        row_key = m.group(2)
        if row_key not in self.rows: return None
        
        row_obj = self.rows[row_key]
        if coordinate not in row_obj.cells: return None
        
        cell_obj = row_obj.cells[ coordinate ]
        return cell_obj

    def create_cell(self, coordinate):
        m = re.match('([A-Z]+)([0-9]+)', coordinate)
        if not m: return False
        
        row_key = m.group(2)
        if row_key not in self.rows: self.create_row(row_key, "")
            
        row_obj = self.rows[row_key]
        if coordinate not in row_obj.cells:
            row_obj.add_cell(coordinate, "", None)

        cell_obj = row_obj.cells[ coordinate ]
        return cell_obj

    def get_cell_string(self, coordinate):
        cell_obj = self.get_cell(coordinate)
        if not cell_obj: return None
        if not cell_obj.is_string(): return None
        qs = self.strings[ int(cell_obj.val) ]
        return qs.pure
        
    def set_cell_value(self, coordinate, val_str, val_pre=False):
        cell_obj = self.get_cell(coordinate)
        if cell_obj == False: return False
        if cell_obj == None: cell_obj = self.create_cell(coordinate)
            
        z_val = val_str
        z_type = None
        qs = None
        
        if type(val_str) in [str, unicode]:
            if val_str not in self.strings_index:
                qs = QuickString( len(self.strings), val_str = val_str, val_pre = val_pre)
                self.strings.append( qs )
                self.strings_index[qs.pure] = qs 
            else:
                qs = self.strings_index[ val_str ]
            z_val = qs.id
            z_type = 's'
            
        if cell_obj.is_string():
            qs_old = self.strings[ int(cell_obj.val) ]
            #if qs_old.id == qs.id: return True
            qs_old.del_ref(cell_obj.id)

        if z_val == None:
            cell_obj.val = None
            cell_obj.type = None
        else:
            cell_obj.val = str(z_val)
            cell_obj.type = z_type

        if qs: # z_type == 's'
            qs.add_ref(cell_obj.id)

        return True

    def open_xlsx(self):
        xl_xmls = ['/xl/sharedStrings.xml', '/xl/worksheets/sheet1.xml']
        
        xlsx_fname = 'out/report/[' + self.templ.name + ']_' + self.path + '.xlsx'
        
        sub_dir = 'tmp/[' + self.templ.name + ']_' + self.path
        if not os.path.exists(sub_dir):
            os.makedirs(sub_dir)

        self.f_strings = sub_dir + xl_xmls[0]
        self.f_sheet = sub_dir + xl_xmls[1]
            
        if os.path.exists( xlsx_fname ):
            print "Using modified file...", xlsx_fname
            z = zipfile.ZipFile(xlsx_fname, 'r', zipfile.ZIP_DEFLATED)
            #for f in xl_xmls: z.extract(f.lstrip('/'), sub_dir)
            z.extractall(sub_dir)
            z.close()
        else:
            print "DBG: Using template file, extract to...", sub_dir
            self.templ.extract( sub_dir )
            #for f in xl_xmls:
            #    full = sub_dir + f
            #    d = full.rsplit('/', 1)
            #    if not os.path.exists(d[0]): os.makedirs(d[0])
            #    shutil.copyfile('tmp/' + self.templ.name + f, full)
                      
    def save_xlsx(self):
        xlsx_fname = 'out/report/[' + self.templ.name + ']_' + self.path + '.xlsx'

        if not os.path.exists('out/report'):
            os.makedirs('out/report')

        print "Writing", xlsx_fname
        
        z = zipfile.ZipFile(xlsx_fname, 'w', zipfile.ZIP_DEFLATED)

        sub_dir = 'tmp/[' + self.templ.name + ']_' + self.path
        files = listall( sub_dir )
        
        for f in files:
            name = f.split('/', 2)[-1]
            z.write(f, name)
        
        #z.write(self.f_strings, 'xl/sharedStrings.xml')
        #z.write(self.f_sheet, 'xl/worksheets/sheet1.xml')
        z.close()

    def save(self):
        self.save_sheet()
        self.save_strings()
        self.save_xlsx()

# u'\u2714'
# u'\xe2\x9c\x94' non utf-8
checkmark = u'\u2714' 

if __name__ == "__main__":
	# Testing purpose only
    MODE = 1
    if MODE == 0:
        qe = QuickExcel('cobacoba', 'opticalport')
        
        print qe.rows_key
        c_a4 = qe.get_cell('A4')
        qe.set_cell_value('A5', 2500)
        c_a5 = qe.get_cell('A5')
        c_a5.style = c_a4.style
        c_j4 = qe.get_cell('J4')
        qe.set_cell_value('J5', "tauco\nmakan\nayam goreng", True)
        c_j5 = qe.get_cell('J5')
        c_j5.style = c_j4.style
        qe.set_cell_value('J6', "tumis\nmakan\nayam bakar", True)
        
        #qe.save_sheet()
        #qe.save_strings()
    elif MODE == 1:
        #extract_template('checklist')
        qe = QuickExcel('cobacoba', 'checklist')
        qe.set_cell_value("I3", "JKT-ESS7")
        qe.set_cell_value("I4", "7450 ESS7")
        qe.set_cell_value("AD39", checkmark)
        
        qe.set_cell_value("AH26", checkmark)
        qe.set_cell_value("AH21", checkmark)
        qe.set_cell_value("AH20", checkmark)

        qe.set_cell_value("L35", 4000000000)
        
        qe.save_sheet()
        qe.save_strings()
        qe.save_xlsx()
    elif MODE == 2:
        qe = QuickExcel('cobacoba', 'checklist')
