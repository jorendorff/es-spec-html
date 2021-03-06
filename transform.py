from xml.etree import ElementTree
from htmodel import *
from docx import shorten, parse_pr

def dict_to_css(d):
    return "; ".join(p + ": " + v for p, v in d.items())

# If True, allow <w:delText> and <w:delInstrText>, ignoring them.
ALLOW_CHANGES = True


# === main

def transform(docx):
    return transform_element(docx, docx.document)

def is_deleted(element, pr_child_name):
    for pr in element:
        if shorten(pr.tag) == pr_child_name:
            for j in pr:
                if shorten(j.tag) == 'del':
                    return True
            return False
    return False

def transform_element(docx, e):
    name = shorten(e.tag)
    assert e.tail is None

    if name == 't':
        assert len(e) == 0
        return e.text

    elif name == 'instrText':
        assert len(e) == 0
        # To translate the Intl spec correctly would involve finding sequences
        # like:
        #
        #     <r><fldChar fldCharType="begin" /></r>
        #     <r><instrText> REF _Ref277198209 \h </instrText></r>
        #     <r><fldChar fldCharType="separate" />
        #     ...
        #     <r><fldChar fldCharType="end" />
        #
        # This might have other benefits, too, like making the table of
        # contents easier to find and making us less dependent upon the author
        # to remember to update fields before saving.
        # 
        if e.text.startswith(' REF '):
            # The REF field: https://office.microsoft.com/en-us/word-help/field-codes-ref-field-HA102017423.aspx
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

    elif name == 'ins':
        assert ALLOW_CHANGES
        return [transform_element(docx, k) for k in e]

    elif name in ('del', 'delText', 'delInstrText', 'moveFrom'):
        assert ALLOW_CHANGES
        return None

    elif name == 'compat:AlternateContent':
        assert shorten(e[0].tag) == 'compat:Choice'
        return transform_element(docx, e[0])

    elif name == 'pic:pic':
        # DrawingML Pictures - http://officeopenxml.com/drwPic.php
        # The actual image is given by e/pic:blipFill/a:blip/@r:embed
        # and the file word/_rels/document.xml.rels in the docx zip.
        image = img()
        for k in e:
            if shorten(k.tag) == 'pic:nvPicPr':  # "non-visual picture properties"
                for gk in k:
                    if shorten(k.tag) == 'pic:cNvPr':  # no idea
                        image.attrs['title'] = gk.get("name", '?')
        return image
    else:
        assert e.text is None

        # Transform all children.
        css = {}
        c = []
        def last_is_deleted():
            if len(c) == 0:
                return False
            last = c[-1]
            return (isinstance(last, Element)
                    and last.name == 'p'
                    and last.style is not None
                    and last.style.get('-ooxml-deleted') == '1')

        def add(ht):
            if isinstance(ht, dict):
                css.update(ht)
            elif isinstance(ht, list):
                for item in ht:
                    add(item)
            elif isinstance(ht, str) and c and isinstance(c[-1], str):
                # Merge adjacent strings.
                c[-1] += ht
            elif (isinstance(ht, Element)
                  and c
                  and isinstance(c[-1], Element)
                  and last_is_deleted()):
                # Merge paragraphs that were joined by deleting the paragraph break.
                #print("Merging this:\n" + repr(c[-1]) + "into this:\n" + repr(ht))
                if ht.name == 'p':
                    c[-1] = ht.with_content(c[-1].content + ht.content)
                else:
                    del c[-1]
                    c.append(ht)
            elif ht is not None:
                c.append(ht)

        for k in e:
            add(transform_element(docx, k))

        if last_is_deleted():
            del c[-1]

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
            result = p(*c)
            if css and '@cls' in css:
                cls = css.pop('@cls')
            else:
                cls = 'Normal'
            result.attrs['class'] = cls
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
            if len(c) == 0:
                return None
            tbl = table(*c)
            ##tbl.style = css
            return figure(tbl)

        elif name == 'tr':
            if is_deleted(e, 'trPr'):
                return None
            return tr(*c)

        elif name == 'tc':
            if is_deleted(e, 'tcPr'):
                return None
            result = td(*c)
            result.style = css
            return result

        else:
            return c

__all__ = ['transform', 'shorten']

