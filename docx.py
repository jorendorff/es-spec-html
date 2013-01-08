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
k_firstLine = bloat('firstLine')
k_before = bloat('before')
k_after = bloat('after')
k_line = bloat('line')

def parse_color(s):
    if s is not None and re.match(r'^[0-9a-fA-F]{6}$', s):
        return '#' + s
    else:
        return None

def parse_borders(k):
    for side in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
        for side_style in k.findall(bloat(side)):
            if side_style.get(k_val) == 'single':
                color = parse_color(side_style.get(k_color)) or 'black'
                sz = side_style.get(k_sz)
                if sz is not None:
                    sz = int(sz)
                    if sz % 6 == 0:
                        sz = str(sz // 6)
                    else:
                        sz = format(sz / 6, '.2f')
                prop = 'border-' + side
                if side.startswith('inside'):
                    prop = '-ooxml-' + prop
                yield (prop, '{}px solid {}'.format(sz, color))

def twips(n):
    if n % 20 == 0:
        return '{}pt'.format(n // 20)
    else:
        return '{:.2f}pt'.format(n / 20)

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
            v = k.get(k_val, '1')
            if v == '1':
                put('font-style', 'italic')
            elif v == '0':
                put('font-style', 'normal')

        elif name == 'b':
            v = k.get(k_val, '1')
            if v == '1':
                put('font-weight', 'bold')
            elif v == '0':
                put('font-weight', 'normal')

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

                ## if font == 'Symbol':
                ##     font = None  # vomit(); vomit(); forget()
                ## elif font == 'Mistral':
                ##     font = None  # fanciful, drop it
                ## elif font in ('Courier', 'Courier New'):
                ##     font = 'monospace'
                ## elif font in ('Arial', 'Helvetica'):
                ##     font = 'sans-serif'
                ## elif font == 'CG Times':
                ##     font = 'Times New Roman'

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

        elif name in ('pBdr', 'tcBorders', 'tblBorders'):
            # TODO w:pBdr>w:between

            # tblBorders can have insideH/insideV elements that are applied to
            # all horizontal/vertical borders between cells in the table. For
            # now, we store that style information in CSS properties named
            # -ooxml-border-insideH/insideV; later we will turn that into
            # border-top/left properties on all the individual table cells.
            for prop, val in parse_borders(k):
                put(prop, val)

        elif name == 'sz':
            if list(k.keys()) == [k_val]:
                # The unit of w:sz is double-points lol.
                v = float(k.get(k_val)) / 2
                put('font-size', str(v) + 'pt')

        elif name == 'jc':
            if list(k.keys()) == [k_val]:
                v = k.get(k_val)
                if v in ('left', 'right', 'center'):
                    put('text-align', v)
                elif v == 'both':
                    put('text-align', 'justify')

        elif name == 'spacing':
            # details of w:spacing differ from CSS: spacing between paragraphs is never
            # less than the computed line-height of either paragraph... or something.
            before = k.get(k_before)
            if before is not None:
                put('margin-top', twips(int(before)))
            after = k.get(k_after)
            if after is not None:
                put('margin-bottom', twips(int(after)))

            # The "1.2 *" below is a heuristic hack.
            line = k.get(k_line)
            if line is not None:
                put('line-height', '{:.1%}'.format(1.2 * float(line) / 240))

        # todo: contextualSpacing

        elif name == 'ind':
            left = k.get(k_left)
            if left is not None:
                left = int(left)
                assert left >= 0
                put('margin-left', twips(left))

                hanging = k.get(k_hanging)
                if hanging is not None:
                    hanging = int(hanging)
                    assert hanging >= 0
                    put('text-indent', twips(-hanging))
                else:
                    firstLine = k.get(k_firstLine)
                    if firstLine is not None:
                        put('text-indent', twips(int(firstLine)))

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
            if shorten(e.tag) == 'pPr':
                # This rPr actually applies to the pilcrow symbol that can
                # (optionally) be displayed at the end of the paragraph.
                # Disdain it.
                pass
            else:
                for k, v in parse_pr(k).items():
                    # TODO - support these properly
                    if k == 'background-color' or k == '@cls':
                        continue
                    put(k, v)

    if pr.get('vertical-align') in ('super', 'sub'):
        sz = pr.get('font-size')
        if sz is None:
            sz = '70%'
        else:
            n, u = re.match(r'([0-9.]*)(.*)', sz).groups()
            sz = format(float(n) * 0.70, '.2f') + u
        pr['font-size'] = sz

    return pr

k_style = bloat('style')
k_styleId = bloat('styleId')

class Style:
    def __init__(self, id, basedOn, type):
        self.id = id
        self.basedOn = basedOn
        self.style = {}
        self.full_style = None
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
k_lvlText = bloat('lvlText')
k_num = bloat('num')
k_numFmt = bloat('numFmt')
k_numId = bloat('numId')
k_numStyleLink = bloat('numStyleLink')
k_numbering = bloat('numbering')
k_pStyle = bloat('pStyle')
k_startOverride = bloat('startOverride')
k_sz = bloat('sz')

class Num:
    """ Data from a <w:num> element. """
    def __init__(self, abstract_num_id, overrides):
        self.abstract_num_id = abstract_num_id
        self.overrides = overrides

class Lvl:
    """ Data from a <w:lvl> element.

    self.pStype is str, self.numFmt is str
    self.lvlText is str
    self.suff is str.
    """
    def render_list_marker(self, numbers):
        # If there's a list-style-type, CSS generates the marker. No need to put one in the HTML.
        if "list-style-type" in self.full_style:
            return None

        def repl(m):
            ilvl = int(m.group(1)) - 1
            return str(numbers[ilvl])  # This really should take into account numFmt, unless isLgl.
        text = re.sub(r'%([1-9])', repl, self.lvlText)
        return text + self.suff

def get_val(e, key, defaultVal=None):
    kids = list(e.findall(bloat(key)))
    if kids:
        [kid] = kids
        val = kid.get(k_val)
        assert val is not None
        return val
    else:
        return defaultVal

list_styles = {
    "none": "none",
    "bullet": "disc",
    "lowerLetter": "lower-alpha",
    "lowerRoman": "lower-roman",
    "upperLetter": "upper-alpha",
    "upperRoman": "upper-roman",
    "decimal": "decimal"
}

suff_chars = {
    'tab': "\t",
    'space': " ",
    'nothing': ""
}

def parse_lvl(docx, e):
    lvl = Lvl()
    assert e.tag == k_lvl
    lvl.pStyle = get_val(e, 'pStyle')
    lvl.numFmt = get_val(e, 'numFmt')
    lvl.lvlText = get_val(e, 'lvlText')
    lvl.suff = suff_chars[get_val(e, 'suff', 'tab')]
    # <w:lvlRestart> also affects numbering in some strange cases.
    # <w:isLgl> affects numbering when present.
    if lvl.pStyle is None:
        style = {}
    else:
        style = docx.styles[lvl.pStyle].full_style.copy()

    for kid in e.findall(bloat('pPr')):
        style.update(parse_pr(kid))
    for kid in e.findall(bloat('rPr')):
        style.update(parse_pr(kid))

    # If this kind of list numbering can be done using simple CSS, do that
    if lvl.suff == '\t':
        ilvl = int(e.get(k_ilvl))
        if lvl.lvlText == '\uf0b7':
            style['list-style-type'] = 'disc'
        elif lvl.numFmt in list_styles and lvl.lvlText == '%{}.'.format(ilvl + 1):
            style['list-style-type'] = list_styles[lvl.numFmt]

        # If either of the two rules above matched...
        if 'list-style-type' in style:
            # then make this a list-item, and remove the text-indent since that
            # is meant to position the marker.
            style['display'] = 'list-item'
            if 'text-indent' in style:
                del style['text-indent']

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
    """ Numbering style data.
    self.abstract_num is {str: [Lvl]}
    self.num is {str: Num}
    """
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
        save(self.styles_xml, 'styles.xml')
        save(self.numbering_xml, 'numbering.xml')

    def _dump_styles(self):
        print(self._style_css())

    def _style_css(self):
        rules = ["p {\n    margin: 0;\n}\n",
                 ".real-table {\n    border-collapse: collapse;\n}\n"]

        def add_rule(selector, props):
            rule = selector + " {\n"
            for prop, value in sorted(props.items()):
                rule += "    " + prop + ": " + value + ";\n"
            rule += "}\n"
            rules.append(rule)

        # Paragraph and character styles
        for cls, s in sorted(self.styles.items()):
            tagname = 'p'
            if s.type == 'character':
                tagname = 'span'
            selector = tagname + "." + cls
            add_rule(selector, s.full_style)

        # Numbering styles
        # Since these follow the paragraph styles, and they have the same CSS
        # specificity, numbering styles take precedence.
        for abstract_num_id, abstract_num in self.numbering.abstract_num.items():
            if not isinstance(abstract_num, str):
                assert isinstance(abstract_num, list)
                for ilvl, lvl in enumerate(abstract_num):
                    if lvl is not None:
                        add_rule("p.abstractnumid-{}-ilvl-{}".format(abstract_num_id, ilvl), lvl.full_style)
        for numid, num in self.numbering.num.items():
            for ilvl, ov in enumerate(num.overrides):
                if ov is not None:
                    add_rule("p.numid-{}-override-ilvl-{}".format(numid, ilvl), ov.full_style)

        return '\n'.join(rules)

    def get_list_class_and_marker_at(self, numId, num_context):
        """ Return the CSS class name corresponding to the given numId and ilvl. """
        num = self.numbering.num[numId]
        ilvl = len(num_context) - 1
        ovs = num.overrides
        if ilvl < len(ovs) and ovs[ilvl] is not None:
            ov = ovs[ilvl]
            cls = "numid-{}-override-ilvl-{}".format(numId, ilvl)
            marker = ov.render_list_marker(num_context)
            return cls, marker
        abstract_num = self.numbering.abstract_num[num.abstract_num_id]
        if isinstance(abstract_num, str):
            style = self.styles[abstract_num]
            return self.get_list_class_and_marker_at(style.full_style['-ooxml-numId'], num_context)
        else:
            assert isinstance(abstract_num, list)
            cls = "abstractnumid-{}-ilvl-{}".format(num.abstract_num_id, ilvl)
            marker = abstract_num[ilvl].render_list_marker(num_context)
            return cls, marker

    def get_list_style_at(self, numId, ilvl):
        num = self.numbering.num[numId]
        ilvl = int(ilvl)
        ov = num.overrides
        if ilvl < len(ov) and ov[ilvl] is not None:
            return ov[ilvl]
        abstract_num = self.numbering.abstract_num[num.abstract_num_id]
        if isinstance(abstract_num, str):
            style = self.styles[abstract_num]
            return self.get_list_style_at(style.full_style['-ooxml-numId'], ilvl)
        else:
            assert isinstance(abstract_num, list)
            return abstract_num[ilvl]

def load(filename):
    with zipfile.ZipFile(filename) as f:
        document_xml = f.read('word/document.xml')
        styles_xml = f.read('word/styles.xml')
        numbering_xml = f.read('word/numbering.xml')

    doc = Document()
    doc.filename = filename
    doc.document = ElementTree.fromstring(document_xml)
    doc.styles_xml = ElementTree.fromstring(styles_xml)
    doc.styles = parse_styles(doc.styles_xml)
    doc.numbering_xml = ElementTree.fromstring(numbering_xml)
    doc.numbering = parse_numbering(doc, doc.numbering_xml)
    return doc
