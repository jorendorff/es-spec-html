import zipfile
from xml.etree import ElementTree
from cgi import escape
import re

namespaces = {
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main': '',
    'http://schemas.openxmlformats.org/markup-compatibility/2006': 'compat',
    'urn:schemas-microsoft-com:vml': 'vml',
    'urn:schemas-microsoft-com:office:office': 'office',
    'http://www.w3.org/XML/1998/namespace': 'xml',
    'urn:schemas-microsoft-com:office:word': 'msword',
}

def shorten(name):
    if name[:1] == '{':
        end = name.index('}')
        schema = name[1:end]
        v = namespaces.get(schema)
        if v is None:
            return name
        elif v == '':
            return name[end + 1:]
        else:
            return v + ':' + name[end + 1:]
    else:
        return name

def bloat(name):
    assert ':' not in name
    return '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + name

k_val = bloat("val")
k_ascii = bloat('ascii')
k_hAnsi = bloat('hAnsi')
k_cs = bloat('cs')
k_eastAsia = bloat('eastAsia')
k_fill = bloat('fill')
k_color = bloat('color')
k_left = bloat('left')
k_hanging = bloat('hanging')


def parse_color(s):
    if s is not None and re.match(r'^[0-9a-fA-F]{6}$', s):
        return '#' + s
    else:
        return None

def parse_pr(e):
    font_keys = {k_ascii, k_hAnsi, k_cs, k_eastAsia, bloat('hint')}

    assert e.text is None

    pr = {}
    def put(k, v):
        if k in pr and pr[k] != v:
            raise Exception("duplicate CSS property on the same element: " + k)
        pr[k] = v

    for k in e:
        assert k.tail is None
        name = shorten(k.tag)

        # TODO: caps, smallCaps, vanish, u, sz

        if name == 'i':
            if not k.keys():
                put('font-style', 'italic')

        elif name == 'b':
            if not k.keys():
                put('font-weight', 'bold')

        elif name == 'rFonts':
            # There are some cases of <rFonts eastAsia="Calibri"/> which Word
            # apparently ignores; we ignore them too.  When cs= and ascii= are
            # both present with different values (awesome), apparently Word
            # uses ascii=, so we do the same.
            assert set(k.keys()) <= font_keys
            font = k.get(k_ascii) or k.get(k_cs)
            if font is not None:
                assert k.get(k_ascii, font) == font
                assert k.get(k_hAnsi, font) == font

                if font == 'Symbol':
                    font = None  # vomit(); vomit(); forget()
                elif font == 'Mistral':
                    font = None  # fanciful, drop it
                elif font in ('Courier', 'Courier New'):
                    font = 'monospace'
                elif font in ('Arial', 'Helvetica'):
                    font = 'sans-serif'
                elif font == 'CG Times':
                    font = 'Times New Roman'

                if font is not None:
                    put('font-family', font)

        elif name == 'vertAlign':
            if list(k.keys()) == [k_val]:
                v = k.get(k_val)
                if v == 'superscript':
                    put('vertical-align', 'super')
                elif v == 'subscript':
                    put('vertical-align', 'sub')

        elif name == 'shd':
            val = k.get(k_val)
            if val == 'solid':
                color = parse_color(k.get(k_color))
            elif val == 'clear':
                color = parse_color(k.get(k_fill))
            else:
                color = None
                    

            if color is not None:
                put('background-color', color)

        elif name in ('tcBorders', 'tblBorders'):
            # tblBorders can have insideH/insideV elements that are applied to
            # all horizontal/vertical borders between cells in the table. For
            # now, we store that style information in CSS properties named
            # -ooxml-border-insideH/insideV; later we will turn that into
            # border-top/left properties on all the individual table cells.
            for side in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
                for side_style in k.findall(bloat(side)):
                    if side_style.get(k_val) == 'single':
                        color = parse_color(side_style.get(k_color)) or 'black'
                        sz = side_style.get(k_sz)
                        if sz is not None:
                            sz = int(sz) // 6
                        prop = 'border-' + side
                        if side.startswith('inside'):
                            prop = '@' + prop
                        put('border-' + side, '{}px solid {}'.format(sz, color))

        elif name == 'sz':
            if list(k.keys()) == [k_val]:
                # The unit of w:sz is half-points lol.
                v = float(k.get(k_val)) / 2
                #put('font-size', str(v) + 'pt')

        # todo: jc, spacing, contextualSpacing
        # todo: pBdr

        elif name == 'ind':
            left = k.get(k_left)
            if left is not None:
                left = int(left)
                assert left >= 0
                #put('padding-left', str(left / 20) + 'pt')
        
                hanging = k.get(k_hanging)
                if hanging is not None:
                    hanging = int(hanging)
                    assert hanging >= 0
                    #put('text-indent', str((hanging - left) / 20) + 'pt')

        elif name == 'numPr':
            ilvl = None
            numId = None
            for item in k:
                item_tag = shorten(item.tag)
                if item_tag == 'ilvl':
                    assert list(item.keys()) == [k_val]
                    assert ilvl is None
                    ilvl = item.get(k_val)
                elif item_tag == 'numId':
                    assert list(item.keys()) == [k_val]
                    assert numId is None
                    numId = item.get(k_val)

            if numId is not None and numId != "0":
                put('-ooxml-numId', numId)
                if ilvl is not None:
                    put('-ooxml-ilvl', ilvl)

        elif name in ('pStyle', 'rStyle'):
            if list(k.keys()) == [k_val]:
                put('@cls', k.get(k_val))

        elif name == 'rPr':
            for k, v in parse_pr(k).items():
                # TODO - support these properly
                if k == 'background-color' or k == '@cls':
                    continue
                put(k, v)

    return pr

k_style = bloat('style')
k_styleId = bloat('styleId')

class Style:
    def __init__(self, id, basedOn, type):
        self.id = id
        self.basedOn = basedOn
        self.style = {}
        self.full_style = None
        assert type in ('paragraph', 'character', 'table', 'numbering')
        self.type = type

k_basedOn = bloat('basedOn')
k_type = bloat('type')
k_pPr = bloat('pPr')
k_rPr = bloat('rPr')

def parse_style(e):
    assert e.tag == k_style
    basedOn_elt = e.find(k_basedOn)
    if basedOn_elt is None:
        basedOn = None
    else:
        basedOn = basedOn_elt.get(k_val)
    s = Style(e.get(k_styleId), basedOn, type=e.get(k_type))

    pPr = e.find(k_pPr)
    if pPr is not None:
        s.style.update(parse_pr(pPr))
    rPr = e.find(k_rPr)
    if rPr is not None:
        s.style.update(parse_pr(rPr))
    return s

def parse_styles(e):
    assert e.tag == bloat('styles')

    all_styles = {}
    for k in e.findall(k_style):
        s = parse_style(k)
        assert s.id not in all_styles
        all_styles[s.id] = s

    def populate_full_style(s):
        if s.full_style is None:
            if s.basedOn is None:
                s.full_style = s.style
            else:
                parent = all_styles[s.basedOn]
                populate_full_style(parent)
                s.full_style = parent.full_style.copy()
                s.full_style.update(s.style)

    for s in all_styles.values():
        populate_full_style(s)

    return all_styles

k_abstractNum = bloat('abstractNum')
k_abstractNumId = bloat('abstractNumId')
k_ilvl = bloat('ilvl')
k_lvl = bloat('lvl')
k_lvlOverride = bloat('lvlOverride')
k_num = bloat('num')
k_numFmt = bloat('numFmt')
k_numId = bloat('numId')
k_numStyleLink = bloat('numStyleLink')
k_numbering = bloat('numbering')
k_pStyle = bloat('pStyle')
k_startOverride = bloat('startOverride')
k_sz = bloat('sz')

class Num:
    def __init__(self, abstract_num_id, overrides):
        self.abstract_num_id = abstract_num_id
        self.overrides = overrides

class Lvl:
    pass

def get_val(e, key):
    kids = list(e.findall(bloat(key)))
    if kids:
        [kid] = kids
        return kid.get(k_val)
    else:
        return None

def parse_lvl(docx, e):
    lvl = Lvl()
    assert e.tag == k_lvl
    lvl.pStyle = get_val(e, 'pStyle')
    lvl.numFmt = get_val(e, 'numFmt')

    if lvl.pStyle is None:
        style = {}
    else:
        style = docx.styles[lvl.pStyle].full_style.copy()
    for kid in e.findall('pPr'):
        style.update(parse_pr(kid))
    for kid in e.findall('rPr'):
        style.update(parse_pr(kid))
    lvl.full_style = style

    return lvl

class StartOverride:
    pass

def parse_startOverride(e):
    assert e.tag == k_startOverride
    ov = StartOverride()
    ov.val = int(e.get(k_val))
    return ov

class Numbering:
    def __init__(self, abstract_num, num):
        self.abstract_num = abstract_num
        self.num = num

def parse_numbering(docx, e):
    assert e.tag == k_numbering

    # eat crunchy xml, num num num
    abstract_num = {}
    for style in e.findall(k_abstractNum):
        abstract_id = style.get(k_abstractNumId)
        assert abstract_id is not None

        nsl = list(style.findall(k_numStyleLink))
        if len(nsl) == 0:
            levels = []
            for level in style.findall(k_lvl):
                ilvl = int(level.get(k_ilvl))
                while len(levels) <= ilvl:
                    levels.append(None)
                levels[ilvl] = parse_lvl(docx, level)
            abstract_num[abstract_id] = levels
        else:
            assert len(nsl) == 1
            assert len(list(style.findall(k_lvl))) == 0
            abstract_num[abstract_id] = nsl[0].get(k_val)

    # Build the num dictionary (extra level of misdirection in OOXML, awesome)
    num = {}
    for style in e.findall(k_num):
        numId = style.get(k_numId)
        assert numId is not None
        val = get_val(style, 'abstractNumId')
        assert val is not None
        overrides = []
        for override in style.findall(k_lvlOverride):
            ilvl = int(override.get(k_ilvl))
            while len(overrides) <= ilvl:
                overrides.append(None)
            [ov] = override
            if ov.tag == k_lvl:
                overrides[ilvl] = parse_lvl(docx, ov)
            else:
                so = parse_startOverride(ov)
                assert so.val == 1
        num[numId] = Num(val, overrides)

    return Numbering(abstract_num, num)

class Document:
    def _extract(self):
        def writexml(e, out, indent='', context='block'):
            t = shorten(e.tag)
            assert e.tail is None
            start_tag = t
            for k, v in e.items():
                start_tag += ' {0}="{1}"'.format(shorten(k), escape(v, True))

            kids = list(e)
            if kids:
                assert e.text is None
                out.write("{0}<{1}>\n".format(indent, start_tag))
                for k in kids:
                    writexml(k, out, indent + '  ')
                out.write("{0}</{1}>\n".format(indent, t))
            elif e.text:
                out.write("{0}<{1}>{2}</{3}>\n".format(indent, start_tag, escape(e.text), t))
            else:
                out.write("{0}<{1} />\n".format(indent, start_tag))

        def save(tree, filename):
            with open(filename, 'w', encoding='utf-8') as out:
                writexml(tree, out)

        save(self.document, 'original.xml')
        #save(self.styles, 'styles.xml')
        #save(self.numbering, 'numbering.xml')

    def _dump_styles(self):
        for cls, s in sorted(self.styles.items()):
            tagname = 'p'
            if s.type == 'character':
                tagname = 'span'
            print(tagname + "." + cls + " {")
            for prop, value in s.style.items():
                print("    " + prop + ": " + value + ";")
            if s.basedOn is not None:
                parent = self.styles[s.basedOn]
                for prop, value in parent.full_style.items():
                    if prop not in s.style:
                        print("    " + prop + ": " + value + ";  /* inherited */")
            print("}\n")

    def get_list_style_at_level(self, numId, ilvl):
        num = self.numbering.num[numId]
        ilvl = int(ilvl)
        ov = num.overrides
        if ilvl < len(ov) and ov[ilvl] is not None:
            return ov[ilvl]
        abstract_num = self.numbering.abstract_num[num.abstract_num_id]
        if isinstance(abstract_num, str):
            style = self.styles[abstract_num]
            return self.get_list_style_at_level(style.full_style['-ooxml-numId'], ilvl)
        else:
            assert isinstance(abstract_num, list)
            if ilvl >= len(abstract_num):
                return None
            else:
                return abstract_num[ilvl]

def load(filename):
    with zipfile.ZipFile(filename) as f:
        document_xml = f.read('word/document.xml')
        styles_xml = f.read('word/styles.xml')
        numbering_xml = f.read('word/numbering.xml')

    doc = Document()
    doc.filename = filename
    doc.document = ElementTree.fromstring(document_xml)
    doc.styles = parse_styles(ElementTree.fromstring(styles_xml))
    doc.numbering = parse_numbering(doc, ElementTree.fromstring(numbering_xml))
    return doc
