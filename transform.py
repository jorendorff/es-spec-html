import collections
from xml.etree import ElementTree
from htmodel import *
from docx import shorten, parse_pr

def dict_to_css(d):
    return "; ".join(p + ": " + v for p, v in d.items())

unrecognized_styles = collections.defaultdict(int)

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
        # Presentation data which we currently do not parse.
        # tcPr has shading which we'd like to have.
        return parse_pr(e)

    elif name == 'pPrChange':
        # A diff to a previous version of the document.
        return None

    else:
        assert e.text is None

        # Transform all children.
        css = {}
        c = []
        def add(ht):
            if isinstance(ht, dict):
                css.update(ht)
            elif isinstance(ht, list):
                for item in ht:
                    add(item)
            elif isinstance(ht, str) and c and isinstance(c[-1], str):
                # Merge adjacent strings.
                c[-1] += ht
            elif ht is not None:
                c.append(ht)

        for k in e:
            add(transform(k))
        if not css:
            css = None

        if name == 'document':
            [body_e] = c
            return html(
                head(
                    title("ECMAScript Language Specification ECMA-262 6th Edition - DRAFT"),
                    link(rel="stylesheet", type="text/css", href="es6-draft.css")),
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
            if css is None:
                return c
            else:
                # No amount of style matters if there's no text here.
                if len(c) == 0:
                    return None
                elif len(c) == 1 and isinstance(c[0], str) and c[0].strip() == '':
                    return c[0] or None

                result = span(*c)
                result.style = css
                return result

        elif name == 'p':
            if len(c) == 0:
                return None
            elif css is None:
                return p(*c)
            else:
                num = '-ooxml-numId' in css
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
                        elif cls in ('Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5',
                                     'Introduction', 'TermNum'):
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
                    'F0B4': '\u00d7', # multiplication sign
                    'F0B8': '\u00f7', # division sign
                    'F0B9': '\u2260', # not equal to
                    'F0CF': '\u2208', # element of
                    'F0D4': '\u2122', # trade mark sign
                    'F0E4': '\u2122'  # trade mark sign (again)
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

        elif name == 'lastRenderedPageBreak':
            # This means "the last time we actually rendered this document to
            # pages, there was a page break here". Theoretically, this could be
            # used to show PDF page numbers in the HTML, but it's not worth it.
            # Everyone uses section numbers anyway.
            return None

        elif name == 'noBreakHyphen':
            # This appears 4 times in the document. The first 3 times it is a
            # mistake and U+2212 MINUS SIGN would be more appropriate. The last
            # time, a plain old hyphen would be better.
            return '\u2011'  #non-breaking hyphen

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
