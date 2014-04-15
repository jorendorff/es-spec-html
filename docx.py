import zipfile
from xml.etree import ElementTree
from html import escape
import re
import copy
from warnings import warn

namespaces = {
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main': '',
    'http://schemas.openxmlformats.org/markup-compatibility/2006': 'compat',
    'urn:schemas-microsoft-com:vml': 'vml',
    'urn:schemas-microsoft-com:office:office': 'office',
    'http://www.w3.org/XML/1998/namespace': 'xml',
    'urn:schemas-microsoft-com:office:word': 'msword',
    'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
    'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r'
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

k_val = bloat('val')
k_ascii = bloat('ascii')
k_hAnsi = bloat('hAnsi')
k_cs = bloat('cs')
k_eastAsia = bloat('eastAsia')
k_fill = bloat('fill')
k_color = bloat('color')
k_left = bloat('left')
k_hanging = bloat('hanging')
k_firstLine = bloat('firstLine')


def parse_color(s):
    if s is not None and re.match(r'^[0-9a-fA-F]{6}$', s):
        return '#' + s
    else:
        return None

def parse_pr(e):
    font_keys = {
        k_ascii, bloat('asciiTheme'),
        k_hAnsi, bloat('hAnsiTheme'),
        k_cs, bloat('cstheme'),
        k_eastAsia, bloat('eastAsiaTheme'),
        bloat('hint')
    }

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
            # ascii, hAnsi, cs, and eastAsia are four different fixed subsets
            # of the Unicode character set. A single w:rFonts element can
            # contain all four, and the corresponding run of text is rendered
            # using one of four fonts, for each character, depending on which
            # subset that character falls into.
            #
            # We don't implement any of that, because for any given w:rFonts
            # element, it seems the same font is specified for all the
            # attributes that are actually defined. (There is only one
            # exception in the document; we ignore it.)
            #
            # It's unclear what is supposed to happen when one or more of the
            # four attributes is missing.
            #
            # In addition there are four more possible attributes: asciiTheme,
            # hAnsiTheme, cstheme, eastAsiaTheme.  Handling these correctly
            # apparently requires you to parse a whole extra file, theme1.xml.

            # These are the fonts mentioned in rFonts elements in the Language
            # Specification, including both document.xml and styles.xml:
            #   Arial, Arial Unicode MS, ArialMT, CG Times, Century,
            #   Courier New, Garamond, Geneva, Helvetica, MS Gothic,
            #   Mistral, Symbol, Tahoma, Times, Times New Roman, Tms Rmn,
            #   Verdana, Wingdings

            assert set(k.keys()) <= font_keys
            font = k.get(k_ascii) or k.get(k_cs)
            if font is not None:
                assert k.get(k_ascii, font) == font
                hAnsiFont = k.get(k_hAnsi, font)
                if hAnsiFont != font:
                    warn("rFonts: font is {} but hAnsi='{}'; ignoring hAnsi.".format(font, hAnsiFont))

                if font == 'Symbol':
                    font = None  # appears once in the document, superfluous
                elif font == 'Mistral':
                    font = None  # fanciful, drop it
                elif font == 'Courier New':
                    font = 'monospace'
                elif font in ('Arial', 'ArialMT', 'Arial Unicode MS', 'Helvetica'):
                    font = 'sans-serif'
                elif font in ('CG Times', 'Times', 'Tms Rmn'):
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
            def fetch(key, css_prop, is_flipped=False):
                val = k.get(key)
                if val is not None:
                    val = round(int(val), -1)  # round to nearest ten (nearest half point)
                    assert val >= 0
                    if is_flipped:
                        val = -val
                    put(css_prop, str(val / 20) + 'pt')

            fetch(k_left, 'margin-left')
            fetch(k_firstLine, 'text-indent')
            fetch(k_hanging, 'text-indent', is_flipped=True)

        elif name == 'numPr':
            for item in k:
                item_tag = shorten(item.tag)
                if item_tag == 'ilvl':
                    assert list(item.keys()) == [k_val]
                    put('-ooxml-ilvl', item.get(k_val))
                elif item_tag == 'numId':
                    assert list(item.keys()) == [k_val]
                    put('-ooxml-numId', item.get(k_val))

        elif name in ('pStyle', 'rStyle'):
            if list(k.keys()) == [k_val]:
                put('@cls', k.get(k_val))

        elif name == 'rPr':
            if shorten(e.tag) == 'pPr':
                # This rPr actually applies to the pilcrow symbol that Word can
                # (optionally) display at the end of the paragraph. The only
                # possibly interesting thing here is if the pilcrow is deleted,
                # indicating this paragraph has been joined with the next one.
                if any(shorten(j.tag) == 'del' for j in k):
                    put('-ooxml-deleted', '1')
            else:
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
        self.computed_levels = None

class AbstractNum:
    def __init__(self, levels, num_style_link):
        assert levels is None or num_style_link is None
        self.levels = levels
        self.num_style_link = num_style_link
        self.computed_levels = None

class Lvl:
    """ Data from a <w:lvl> element.

    self.start is int.
    self.pStyle is str.
    self.numFmt is str.
    self.lvlText is str.
    self.suff is str.
    self.style and self.full_style are CSS dictionaries.
    """
    def with_start(self, start):
        clone = copy.copy(self)
        clone.start = start
        return clone

def get_val(e, key, default_value = None):
    kids = list(e.findall(bloat(key)))
    if kids:
        [kid] = kids
        return kid.get(k_val, default_value)
    else:
        return default_value

suff_values = {
    'nothing': '',
    'space': ' ',
    'tab': '\t'
}

def parse_lvl(docx, e):
    lvl = Lvl()
    assert e.tag == k_lvl
    lvl.start = int(get_val(e, 'start', '1'))
    lvl.numFmt = get_val(e, 'numFmt')
    lvl.pStyle = get_val(e, 'pStyle')
    lvl.lvlText = get_val(e, 'lvlText')
    lvl.suff = suff_values[get_val(e, 'suff', 'tab')]

    if lvl.pStyle is None:
        style = {}
    else:
        style = docx.styles[lvl.pStyle].full_style.copy()
    for kid in e.findall(k_pPr):
        style.update(parse_pr(kid))
    for kid in e.findall(k_rPr):
        style.update(parse_pr(kid))
    lvl.full_style = style

    return lvl

class LvlOverride:
    def __init__(self, ilvl, lvl, startOverride):
        self.ilvl = ilvl
        self.lvl = lvl
        self.startOverride = startOverride

def parse_lvlOverride(docx, e):
    ilvl = int(e.get(k_ilvl))

    lvl = None
    startOverride = None

    all_lvl = e.findall(k_lvl)
    if all_lvl:
        [sole_lvl] = all_lvl
        lvl = parse_lvl(docx, sole_lvl)

    all_startOverride = e.findall(k_startOverride)
    if all_startOverride:
        [sole_startOverride] = all_startOverride
        startOverride = int(sole_startOverride.get(k_val))

    if len(e) == 0:
        startOverride = 0  # land of absurd defaults

    return LvlOverride(ilvl, lvl, startOverride)

def parse_startOverride(e):
    assert e.tag == k_startOverride
    val = int(e.get(k_val))
    return val

NUMBERING_LEVELS = 10

class Numbering:
    """ Represents the whole content of numbering.xml (a w:Numbering element).
    self.abstract_num is {abstract_num_id: AbstractNum object}.
    self.num is {num_id: Num object}.
    self.style_links is {styleLink value: abstract_num_id}.
    """
    def __init__(self, abstract_num, num, style_links):
        self.abstract_num = abstract_num
        self.num = num
        self.style_links = style_links
        for abstract_num_id in self.abstract_num:
            self._compute_abstract_num_levels(abstract_num_id)
        for num in self.num.values():
            abstract_num_id = num.abstract_num_id
            levels = self.abstract_num[abstract_num_id].computed_levels[:]
            for i, lvlOverride in enumerate(num.overrides):
                if lvlOverride is not None:
                    if lvlOverride.lvl is not None:
                        levels[i] = lvlOverride.lvl
                    if lvlOverride.startOverride is not None:
                        levels[i] = levels[i].with_start(lvlOverride.startOverride)
            num.computed_levels = levels

    def _compute_abstract_num_levels(self, abstract_num_id):
        abstract_num = self.abstract_num[abstract_num_id]
        if abstract_num.computed_levels is not None:
            return abstract_num.computed_levels

        levels = []
        for i in range(NUMBERING_LEVELS):
            if abstract_num.num_style_link is not None:
                # i'm sorry mario but the princess
                real_abstract_num = self.style_links[abstract_num.num_style_link]
                base_levels = self._compute_abstract_num_levels(real_abstract_num)
                level = base_levels[i]
            elif i < len(abstract_num.levels):
                level = abstract_num.levels[i]
            else:
                level = None
            levels.append(level)
        abstract_num.computed_levels = levels
        return levels

    def get_abstract_num_id_and_levels(self, numId):
        num = self.num[numId]
        return num.abstract_num_id, num.computed_levels

def parse_numbering(docx, e):
    # See <http://msdn.microsoft.com/en-us/library/ee922775%28office.14%29.aspx>.
    assert e.tag == k_numbering

    # eat crunchy xml, num num num
    abstract_num = {}
    style_links = {}
    for ane in e.findall(k_abstractNum):
        abstract_id = int(ane.get(k_abstractNumId))

        # w:numStyleLink. This is a reference to a w:abstractNum that has a
        # w:styleLink child element.
        nsl = list(ane.findall(k_numStyleLink))
        if len(nsl) == 0:
            levels = []
            for level in ane.findall(k_lvl):
                ilvl = int(level.get(k_ilvl))
                while len(levels) <= ilvl:
                    levels.append(None)
                levels[ilvl] = parse_lvl(docx, level)
            abstract_num[abstract_id] = AbstractNum(levels, None)
        else:
            assert len(nsl) == 1
            assert len(list(ane.findall(k_lvl))) == 0
            abstract_num[abstract_id] = AbstractNum(None, nsl[0].get(k_val))

        # w:styleLink.
        link = get_val(ane, 'styleLink')
        if link is not None:
            assert link not in style_links
            style_links[link] = abstract_id

    # Build the num dictionary (extra level of misdirection in OOXML, awesome)
    num = {}
    for style in e.findall(k_num):
        numId = int(style.get(k_numId))
        val = int(get_val(style, 'abstractNumId'))
        overrides = []
        for override_element in style.findall(k_lvlOverride):
            lvlOverride = parse_lvlOverride(docx, override_element)
            while len(overrides) <= lvlOverride.ilvl:
                overrides.append(None)
            overrides[lvlOverride.ilvl] = lvlOverride
        num[numId] = Num(val, overrides)

    return Numbering(abstract_num, num, style_links)

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
        save(self.styles_raw, 'styles.xml')
        save(self.numbering_raw, 'numbering.xml')

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
        """ Returns a Lvl object; its .full_style attribute is a CSS dictionary. """
        return self.numbering.num[numId].computed_levels[int(ilvl)]

def load(filename):
    with zipfile.ZipFile(filename) as f:
        document_xml = f.read('word/document.xml')
        styles_xml = f.read('word/styles.xml')
        numbering_xml = f.read('word/numbering.xml')

    doc = Document()
    doc.filename = filename
    doc.document = ElementTree.fromstring(document_xml)
    doc.styles_raw = ElementTree.fromstring(styles_xml)
    doc.styles = parse_styles(doc.styles_raw)
    doc.numbering_raw = ElementTree.fromstring(numbering_xml)
    doc.numbering = parse_numbering(doc, doc.numbering_raw)
    return doc
