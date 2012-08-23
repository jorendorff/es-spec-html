from xml.etree import ElementTree
from htmodel import *
from docx import shorten, parse_pr

def dict_to_css(d):
    return "; ".join(p + ": " + v for p, v in d.items())

# If True, allow <w:delText>, ignoring it.
ALLOW_CHANGES = True

def transform(e):
    name = shorten(e.tag)
    assert e.tail is None

    if name == 't':
        assert len(e) == 0
        return e.text

    elif name == 'instrText' or name == 'fldChar':
        assert len(e) == 0
        return None

    elif name in {'pPr', 'rPr', 'sectPr', 'tblPr', 'tblPrEx', 'trPr', 'tcPr', 'numPr'}:
        # Presentation data.
        return parse_pr(e)

    elif name == 'pPrChange':
        # A diff to a previous version of the document.
        return None

    elif name == 'delText':
        assert ALLOW_CHANGES
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
                head(link(rel="stylesheet", href="es6-draft.css")),
                body_e)

        elif name == 'body':
            return body(*c)

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
                if css and '@cls' in css:
                    result.attrs['class'] = css['@cls']
                    del css['@cls']
                return result

        elif name == 'p':
            if len(c) == 0:
                return None

            result = p(*c)
            if css and '@cls' in css:
                result.attrs['class'] = css['@cls']
                del css['@cls']
            result.style = css
            return result

        elif name == 'pict':
            return div(*c, class_='w-pict')

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
                    'F0CF': '\u2209', # not an element of
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
            tbl = table(*c, class_="real-table")
            if css:
                if '-ooxml-border-insideH' in css:
                    # borders between rows
                    row_border = css['-ooxml-border-insideH']
                    del css['-ooxml-border-insideH']
                if '-ooxml-border-insideV' in css:
                    # borders between columns
                    col_border = css['-ooxml-border-insideV']
                    del css['-ooxml-border-insideV']
                ##tbl.style = css
            return figure(tbl)

        elif name == 'tr':
            return tr(*c)

        elif name == 'tc':
            result = td(*c)
            result.style = css
            return result

        else:
            return c

__all__ = ['transform', 'shorten']
