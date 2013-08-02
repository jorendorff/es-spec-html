from xml.etree import ElementTree
from htmodel import *
from docx import shorten, parse_pr
from collections import defaultdict
import re

def dict_to_css(d):
    return "; ".join(p + ": " + v for p, v in d.items())

# If True, allow <w:delText> and <w:delInstrText>, ignoring them.
ALLOW_CHANGES = True

# === numbering

def int_to_lower_roman(i):
    """ Convert an integer to Roman numerals.
    From Paul Winkler's recipe: https://code.activestate.com/recipes/81611-roman-numerals/
    """
    if i < 1 or i > 3999:
        raise ValueError("Argument must be between 1 and 3999")
    vals = (1000, 900,  500, 400, 100,  90, 50,  40, 10,  9,   5,  4,   1)
    syms = ('m',  'cm', 'd', 'cd','c', 'xc','l','xl','x','ix','v','iv','i')
    result = ""
    for val, symbol in zip(vals, syms):
        count = i // val
        result += symbol * count
        i -= val * count
    return result

def int_to_lower_letter(i):
    if i > 26:
        raise ValueError("Don't know any more letters after z.")
    return "abcdefghijklmnopqrstuvwxyz"[i - 1]

list_formatters = {
    'lowerLetter': int_to_lower_letter,
    'upperLetter': lambda i: int_to_lower_letter(i).upper(),
    'decimal': str,
    'lowerRoman': int_to_lower_roman,
    'upperRoman': lambda i: int_to_lower_roman(i).upper()
}

def render_list_marker(levels, numbers):
    def repl(m):
        ilvl = int(m.group(1)) - 1
        level = levels[ilvl]
        return list_formatters[level.numFmt](numbers[ilvl])  # should ignore numFmt if isLgl
    this_level = levels[len(numbers) - 1]
    text = re.sub(r'%([1-9])', repl, this_level.lvlText)
    return text + this_level.suff

# === main

def transform(docx):
    return transform_element(docx, docx.document, numbering_context=defaultdict(list))

def transform_element(docx, e, numbering_context):
    name = shorten(e.tag)
    assert e.tail is None

    if name == 't':
        assert len(e) == 0
        return e.text

    elif name == 'instrText':
        assert len(e) == 0
        if e.text.startswith(' SEQ ') or e.text.startswith(' REF '):
            return '{' + e.text + '}'
        return None

    elif name in {'pPr', 'rPr', 'sectPr', 'tblPr', 'tblPrEx', 'trPr', 'tcPr', 'numPr'}:
        # Presentation data.
        return parse_pr(e)

    elif name == 'pPrChange':
        # A diff to a previous version of the document.
        return None

    elif name in {'{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}posOffset',
                  '{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}pctWidth',
                  '{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}pctHeight'
                 }:
        # Layout data
        return None

    elif name in ('del', 'delText', 'delInstrText'):
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
            add(transform_element(docx, k, numbering_context))
        if not css:
            css = None

        if name == 'document':
            [body_e] = c
            return html(
                head(),
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
                    result.attrs['class'] = css.pop('@cls')
                return result

        elif name == 'p':
            if len(c) == 0:
                return None

            result = p(*c)
            if css and '@cls' in css:
                cls = css.pop('@cls')
            else:
                cls = 'Normal'
            result.attrs['class'] = cls

            # Numbering.
            paragraph_style = docx.styles[cls]
            if css and '-ooxml-numId' in css:
                numid = int(css['-ooxml-numId'])
            elif '-ooxml-numId' in paragraph_style.full_style:
                numid = int(paragraph_style.full_style['-ooxml-numId'])
            else:
                numid = 0

            if numid != 0:
                # Figure out the level of this paragraph.
                if '-ooxml-ilvl' in css:
                    ilvl = int(css['-ooxml-ilvl'])
                elif '-ooxml-ilvl' in paragraph_style.full_style:
                    ilvl = int(paragraph_style.full_style['-ooxml-ilvl'])
                else:
                    ilvl = 0

                # Bump the numbering accordingly.
                abstract_num_id, levels = docx.get_abstract_num_id_and_levels(numid, ilvl)
                current_number = numbering_context[abstract_num_id]
                if len(current_number) <= ilvl:
                    while len(current_number) <= ilvl:
                        start = levels[len(current_number)].start
                        current_number.append(start)
                else:
                    del current_number[ilvl + 1:]
                    current_number[ilvl] += 1

                # Create a suitable marker.
                marker = render_list_marker(levels, current_number)
                s = span(marker, class_="marker")
                s.style = {}
                result.content.insert(0, s)

            result.style = css
            return result

        elif name == 'pict' or name == 'drawing':
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
                    row_border = css.pop('-ooxml-border-insideH')
                if '-ooxml-border-insideV' in css:
                    # borders between columns
                    col_border = css.pop('-ooxml-border-insideV')
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

