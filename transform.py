from xml.etree import ElementTree
from html import *

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

def transformPr(e):
    assert e.text is None

    pr = {}
    def put(k, v):
        if k in pr:
            raise Exception("duplicate CSS property on the same element: " + k)
        pr[k] = v

    for k in e:
        assert k.tail is None
        name = shorten(k.tag)
        if name == 'i':
            if not k.keys():
                put('font-style', 'italic')
        elif name == 'b':
            if not k.keys():
                put('font-weight', 'bold')
        elif name == 'vertAlign':
            if list(k.keys()) == [bloat('val')]:
                v = k.get(bloat('val'))
                if v == 'superscript':
                    put('vertical-align', 'super')
                elif v == 'subscript':
                    put('vertical-align', 'sub')
        elif name == 'pStyle':
            if list(k.keys()) == [bloat('val')]:
                put('@cls', k.get(bloat('val')))

    return pr or None

def dict_to_css(d):
    return "; ".join(p + ": " + v for p, v in d.items())

def transform(e):
    name = shorten(e.tag)
    assert e.tail is None

    if name == 't':
        assert len(e) == 0
        return e.text

    elif name == 'instrText':
        assert len(e) == 0
        return '{' + e.text + '}'

    elif name in {'pPr', 'rPr', 'sectPr', 'tblPr', 'tblPrEx', 'trPr', 'tcPr', 'numPr'}:
        # Presentation data.
        return transformPr(e)

    else:
        assert e.text is None

        # Transform all children.
        css = None
        c = []
        for k in e:
            ht = transform(k)
            if isinstance(ht, dict):
                if css is None:
                    css = ht
                else:
                    css.update(ht)
            elif isinstance(ht, str) and c and isinstance(c[-1], str):
                # Merge adjacent strings.
                c[-1] += ht
            elif ht is not None:
                c.append(ht)

        if name == 'document':
            [body_e] = c
            return html(
                head(
                    meta(http_equiv="Content-Type", content="text/html; charset=UTF-8"),
                    link(rel="stylesheet", type="text/css", href="es-spec.css")),
                body_e)

        elif name == 'body':
            disclaimer = div(
                p(strong("This is ", em("not"), " the official ECMAScript Language Specification.")),
                p("The most recent final ECMAScript standard is Edition 5.1, the PDF document located at ",
                  a("http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf",
                    href="http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf"),
                  "."),
                p("This is a draft of the next version of the standard. If all goes well it will become "
                  "ECMAScript Edition 6."),
                p("This is an HTML version of the current working draft published at ",
                  a("http://wiki.ecmascript.org/doku.php?id=harmony:specification_drafts",
                    href="http://wiki.ecmascript.org/doku.php?id=harmony:specification_drafts"),
                  ". The program used to convert that Word doc to HTML is a custom-piled heap of hacks. "
                  "Currently it is pretty bad. It has stripped out most of the formatting that makes "
                  "the specification comprehensible. You can help improve the program ",
                  a("here", href="https://github.com/jorendorff/es-spec-html"),
                  "."),
                p("For copyright information, see ECMA's legal disclaimer in the document itself."),
                id="unofficial")
            return body(section(disclaimer, *c, id="everything"))

        elif name == 'r':
            if css is not None:
                if css == {'font-style': 'italic'}:
                    return i(*c)
                elif css == {'font-weight': 'bold'}:
                    return b(*c)
                elif css == {'vertical-align': 'super'}:
                    return sup(*c)
                elif css == {'vertical-align': 'sub'}:
                    return sub(*c)
                else:
                    return span(*c, style=dict_to_css(css))
            else:
                if len(c) == 0:
                    return None
                elif len(c) == 1:
                    return c[0]
                else:
                    return span(*c)

        elif name == 'p':
            if css is not None:
                if '@cls' in css:
                    cls = css['@cls']
                    del css['@cls']
                    if not css:
                        if cls == 'Heading1':
                            return h1(*c)
                        elif cls in ('Heading2', 'Heading3', 'Heading4'):
                            return h2(*c)
                        else:
                            return p(*c)
                return p(*c, style=dict_to_css(css))
            else:
                return p(*c)

        elif name == 'sym':
            assert not c
            attrs = {shorten(k): v for k, v in e.items()}
            if len(attrs) == 2 and attrs['font'] == 'Symbol' and 'char' in attrs:
                _symbols = {
                    'F02D': '\u2212', # minus sign
                    'F070': '\u03C0', # greek small letter pi
                    'F0A3': '\u2264', # less-than or equal to
                    'F0A5': '\u221e', # infinity
                    'F0B3': '\u2265', # greater-than or equal to
                    'F0B4': '\u00d7', # times sign
                    'F0B8': '\u00f7', # division sign
                    'F0B9': '\u2260', # not equal to
                    'F0CF': '\u2208', # element of
                    'F0D4': '\u2122', # trademark sign
                    'F0E4': '\u2122'  # trademark sign again
                }
                ch = _symbols.get(attrs['char'], '\ufffd') # U+FFFD, replacement character
                if ch == '\ufffd':
                    ch += ' (' + attrs['char'] + ')'
                    ElementTree.dump(e)
                return ch
            ElementTree.dump(e)
            return None

        elif name == 'tab':
            assert not c
            assert not e.keys()
            return '\t'

        elif name == 'br':
            assert not c
            assert set(e.keys()) <= {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type'}
            br_type = e.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
            if br_type is None:
                return br()
            else:
                assert br_type == 'page'
                return hr()

        elif name in {'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd'}:
            return None

        elif name == 'tbl':
            assert not e.keys()
            return figure(table(*c, class_="real-table"))

        elif name == 'tr':
            return tr(*c)

        elif name == 'tc':
            return td(*c)

        else:
            if len(c) == 0:
                return None
            elif len(c) == 1:
                return c[0]
            else:
                return div(*c)

__all__ = ['transform', 'shorten']
