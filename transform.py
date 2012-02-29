import collections
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
    val_key = bloat('val')
    ascii_key = bloat('ascii')
    hAnsi_key = bloat('hAnsi')
    cs_key = bloat('cs')
    eastAsia_key = bloat('eastAsia')
    font_keys = {ascii_key, hAnsi_key, cs_key, eastAsia_key}

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

        elif name == 'rFonts':
            # There are some cases of <rFonts eastAsia="Calibri"/> which Word
            # apparently ignores; we ignore them too.  When cs= and ascii= are
            # both present with different values (awesome), apparently Word
            # uses ascii=, so we do the same.
            assert set(k.keys()) <= font_keys
            font = k.get(ascii_key) or k.get(cs_key)
            if font is not None:
                assert k.get(ascii_key, font) == font
                assert k.get(hAnsi_key, font) == font

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
                    all_fonts[font] += 1

        elif name == 'vertAlign':
            if list(k.keys()) == [val_key]:
                v = k.get(val_key)
                if v == 'superscript':
                    put('vertical-align', 'super')
                elif v == 'subscript':
                    put('vertical-align', 'sub')

        elif name == 'numPr':
            ilvl = None
            numId = None
            for item in k:
                item_tag = shorten(item.tag)
                if item_tag == 'ilvl':
                    assert list(item.keys()) == [val_key]
                    assert ilvl is None
                    ilvl = item.get(val_key)
                else:
                    assert item_tag == 'numId'
                    assert list(item.keys()) == [val_key]
                    assert numId is None
                    numId = item.get(val_key)

            if numId != "0":
                put('@num', ilvl + '/' + numId)

        elif name == 'pStyle':
            if list(k.keys()) == [val_key]:
                put('@cls', k.get(val_key))

    return pr or None

def dict_to_css(d):
    return "; ".join(p + ": " + v for p, v in d.items())

unrecognized_styles = collections.defaultdict(int)
all_fonts = collections.defaultdict(int)

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
                # No amount of style matters if there's no text here.
                if len(c) == 0:
                    return None
                elif len(c) == 1 and isinstance(c[0], str) and c[0].strip() == '':
                    return c[0] or None

                if css == {'font-style': 'italic'}:
                    return i(*c)
                elif css == {'font-weight': 'bold'}:
                    return b(*c)
                elif css == {'vertical-align': 'super'}:
                    return sup(*c)
                elif css == {'vertical-align': 'sub'}:
                    return sub(*c)
                elif css == {'font-family': 'monospace', 'font-weight': 'bold'}:
                    return code(*c)
                else:
                    result = span(*c)
                    result.style = css
                    return result
            else:
                if len(c) == 0:
                    return None
                elif len(c) == 1:
                    return c[0]
                else:
                    return span(*c)

        elif name == 'p':
            if css is None:
                return p(*c)
            else:
                num = '@num' in css
                constructor = li if num else p

                if '@cls' in css:
                    cls = css['@cls']
                    del css['@cls']
                    if not css:
                        if cls in ('Alg2', 'Alg3', 'Alg4', 'Alg40', 'Alg41', 'M4'):
                            if num or c:
                                return constructor(*c)
                            else:
                                # apparently useless markup
                                return None
                        elif cls == 'ANNEX':
                            # TODO - add annex heading, which is computed in the original
                            return h1(*c)
                        elif cls == 'bibliography':
                            if len(c) == 0:
                                return None
                            return li(*c, class_="bibliography-entry")
                        elif cls == 'BulletNotlast':
                            return li(*c)
                        elif cls in ('CodeSample3', 'CodeSample4'):
                            return pre(*c)
                        elif cls in ('Definition', 'M0'):
                            # apparently useless markup
                            return p(*c)
                        elif cls == 'Figuretitle':
                            return figcaption(*c)
                        elif cls in ('Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'TermNum'):
                            return h1(*c)
                        elif cls == 'M20':
                            return div(*c, class_="math-display")
                        elif cls == 'MathDefinition4':
                            return div(*c, class_="display")
                        elif cls == 'MathSpecialCase3':
                            return li(*c)
                        elif cls == 'Note':
                            return div(constructor(*c), class_="note")
                        elif cls == 'RefNorm':
                            return p(*c, class_="formal-reference")
                        elif cls == 'Syntax':
                            return h2(*c)
                        elif cls in ('SyntaxRule', 'SyntaxRule2'):
                            return div(*c, class_="gp")
                        elif cls in ('SyntaxDefinition', 'SyntaxDefinition2'):
                            return div(*c, class_="rhs")
                        elif cls == 'Tabletitle':
                            return figcaption(*c)
                        elif cls == 'Terms':
                            return p(dfn(*c))
                        elif cls == 'zzBiblio':
                            return h1(*c)
                        elif cls == 'zzSTDTitle':
                            return div(*c, class_="inner-title")
                        else:
                            unrecognized_styles[cls] += 1
                            #return p(span('<{0}>'.format(cls), style="color:red"), *c)
                            return constructor(*c)
                result = constructor(*c)
                result.style = css
                return result

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
