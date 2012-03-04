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

def parse_pr(e):
    font_keys = {k_ascii, k_hAnsi, k_cs, k_eastAsia}

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
            if k.get(k_val) == "solid" and k.get(k_fill) == "FFFFFF":
                color = k.get(k_color)
                if color is not None and re.match(r'^[0-9a-fA-F]{6}$', color):
                    put('background-color', '#' + color)

        # todo: shd, jc, ind, spacing, contextualSpacing
        # todo: pBdr

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
    def __init__(self, id, basedOn):
        self.id = id
        self.basedOn = basedOn
        self.style = {}
        self.full_style = None

k_basedOn = bloat('basedOn')
k_pPr = bloat('pPr')
k_rPr = bloat('rPr')

def parse_style(e):
    assert e.tag == k_style
    basedOn_elt = e.find(k_basedOn)
    if basedOn_elt is None:
        basedOn = None
    else:
        basedOn = basedOn_elt.get(k_val)
    s = Style(e.get(k_styleId), basedOn)

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
        save(self.styles, 'styles.xml')
        save(self.numbering, 'numbering.xml')

    def _dump_styles(self):
        for cls, s in sorted(self.styles.items()):
            print("p." + cls + " {")
            for prop, value in s.style.items():
                print("    " + prop + ": " + value + ";")
            if s.basedOn is not None:
                parent = self.styles[s.basedOn]
                for prop, value in parent.full_style.items():
                    if prop not in s.style:
                        print("    " + prop + ": " + value + ";  /* inherited */")
            print("}\n")

def load(filename):
    with zipfile.ZipFile("es6-draft.docx") as f:
        document_xml = f.read('word/document.xml')
        numbering_xml = f.read('word/numbering.xml')
        styles_xml = f.read('word/styles.xml')

    doc = Document()
    doc.document = ElementTree.fromstring(document_xml)
    doc.numbering = ElementTree.fromstring(numbering_xml)
    doc.styles = parse_styles(ElementTree.fromstring(styles_xml))
    return doc
