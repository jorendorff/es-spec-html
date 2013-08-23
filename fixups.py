""" fixups.py - Refine the HTML document produced by transform.py.

The HTML markup produced by transform.py is extremely crude.
These fixups add links, lists, a stylesheet, and sections.
A great deal of the work done here is document-specific.

The entry point is fixup().
"""

import htmodel as html
import collections, re
from warnings import warn
from array import array
import os, time
import contextlib


# === Useful functions

def findall(e, name):
    if e.name == name:
        yield e
    for k in e.content:
        if not isinstance(k, str):
            for d in findall(k, name):
                yield d

def all_parent_index_child_triples(e):
    for i, k in enumerate(e.content):
        if not isinstance(k, str):
            yield e, i, k
            for t in all_parent_index_child_triples(k):
                yield t

def all_parent_index_child_triples_reversed(e):
    i = len(e.content) - 1
    while i >= 0:
        k = e.content[i]
        if not isinstance(k, str):
            for t in all_parent_index_child_triples_reversed(k):
                yield t
            assert e.content[i] is k
            yield e, i, k
        i -= 1

def spec_is_intl(docx):
    return os.path.basename(docx.filename).lower().startswith('es-intl')

def spec_is_lang(docx):
    return not os.path.basename(docx.filename).lower().startswith('es-intl')

def version_is_5(docx):
    return os.path.basename(docx.filename).lower().startswith('es5')

def version_is_51_final(docx):
    return os.path.basename(docx.filename) == 'es5.1-final.dotx'

def version_is_intl_1_final(docx):
    return os.path.basename(docx.filename) == 'es-intl-1-final.docx'


# === Kinds of fixups

class Fixup:
    def __init__(self, fn):
        self.fn = fn
        self.name = fn.__name__
    def __call__(self, doc, docx):
        result = self.fn(doc, docx)
        assert result is not None
        assert isinstance(result, html.Element)
        return result

class InPlaceFixup:
    """
    A kind of fixup that modifies the document in-place
    rather than creating a copy with some changes.
    """
    def __init__(self, fn):
        self.fn = fn
        self.name = fn.__name__
    def __call__(self, doc, docx):
        result = self.fn(doc, docx)
        assert result is None
        return doc


# === Fixups

@Fixup
def fixup_strip_empty_paragraphs(doc, docx):
    def is_empty_para(e):
        return e.name == 'p' and len(e.content) == 0
    return doc.find_replace(is_empty_para, lambda e: [])

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

@Fixup
def fixup_add_numbering(doc, docx):
    """ Add span.marker elements and -ooxml-indentation style properties to the document. """

    numbering_context = collections.defaultdict(list)

    def add_numbering(p):
        cls = p.attrs and p.attrs.get('class')
        paragraph_style = docx.styles[cls]

        numid = ilvl = None
        def computed_style(name, default_value, using_list_styles=False):
            """
            Get computed style for the given property name.
            Returns a string, or default_value if no such property is defined anywhere.
            """
            # This paragraph's properties override everything else.
            if p.style and name in p.style:
                return p.style[name]

            # Properties inherited from a numbering w:lvl>w:pPr are
            # next-highest in precedence.
            if using_list_styles:
                lvl = docx.get_list_style_at_level(numid, ilvl)
                if lvl is not None and name in lvl.full_style:
                    return lvl.full_style[name]

            # After that come the properties defined in paragraph
            # style. Note that full_style incorporates properties that are
            # inherited via the w:basedOn chain.
            if name in paragraph_style.full_style:
                return paragraph_style.full_style[name]

            # Not specified anywhere.
            return default_value

        numid = int(computed_style('-ooxml-numId', '0'))

        has_numbering = numid != 0
        if has_numbering:
            # Figure out the level of this paragraph.
            ilvl = int(computed_style('-ooxml-ilvl', '0'))

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
            s = html.span(marker, class_="marker")
            s.style = {}
            content = [s] + p.content
        else:
            content = p.content

        # Figure out the actual physical indentation of the number on this
        # paragraph, net of everything. (This is used in fixup_lists_early
        # to infer nesting lists, whether a paragraph is inside a list
        # item, etc.)
        def points(s):
            if s == '0':
                return 0
            assert s.endswith('pt')
            return float(s[:-2])
        margin_left = points(computed_style('margin-left', '0', using_list_styles=has_numbering))
        text_indent = points(computed_style('text-indent', '0', using_list_styles=has_numbering))
        if p.style is None:
            css = {}
        else:
            css = p.style.copy()
        css['-ooxml-indentation'] = str(margin_left + text_indent) + 'pt'

        return p.with_(style=css, content=content)

    def fix_body(body):
        result = []
        for p in body.content:
            if p.name == 'p':
                p = add_numbering(p)
            result.append(p)
        return [body.with_content(result)]

    return doc.find_replace(lambda e: e.name in ('body', 'td'), fix_body)

def has_bullet(docx, p):
    """ True if the given paragraph is of a style that has a bullet. """
    if not p.style:
        return False
    numId = int(p.style.get('-ooxml-numId', '0'))
    if numId == 0:
        return False
    ilvl = p.style.get('-ooxml-ilvl', '0')
    s = docx.get_list_style_at_level(numId, ilvl)
    return s is not None and s.numFmt == 'bullet'

@InPlaceFixup
def fixup_list_styles(doc, docx):
    """ Make sure bullet lists are never p.Alg4 or other particular styles.

    Alg4 style indicates a numbered list, with Times New Roman font. It's used
    for algorithms. However there are a few places in the Word document where
    a paragraph is style Alg4 but manually hacked to have a bullet instead of a
    number and the default font instead of Times New Roman. Lol.

    In short, this screws everything up, so we manually hack it in the general
    direction of sanity before doing anything else.

    Precedes fixup_formatting, which would spew a bunch of
    <span style="font-family: sans-serif"> if we did it first.
    """

    wrong_types = ('Alg4', 'MathSpecialCase3', 'BulletNotlast', 'BulletLast')

    for p in findall(doc, 'p'):
        if p.attrs.get("class") in wrong_types and has_bullet(docx, p):
            p.attrs['class'] = "Normal"

def looks_like_nonterminal(text):
    return re.match(r'^(?:uri(?:[A-Z][A-Za-z]*)?|[A-Z]+[a-z][A-Za-z]*)$', text) is not None

def is_marker(e):
    return ht_name_is(e, 'span') and e.attrs.get('class') == 'marker'

@Fixup
def fixup_formatting(doc, docx):
    """
    Convert runs of span elements to more HTML-like code.

    The OOXML schema starts out like this:
     - w:body contains w:p elements (paragraphs)
     - w:p contains w:r elements (runs)
     - w:r contains an optional w:rPr (style data) followed by some amount of text
       and/or tabs, line breaks, and other junk (w:t, w:tab, w:br, etc.)

    But note that w:r elements don't nest.

    In Word, when someone selects a whole paragraph, already containing markup,
    and changes the font, the result is a paragraph containing many runs, every
    one of which has its own w:r>w:rPr>w:rFonts element. This fixup turns the
    redundant formatting into nested HTML markup.
    """

    def new_span(content, style):
        # Merge adjacent strings, if any.
        i = 0
        while i < len(content) - 1:
            a = content[i]
            b = content[i + 1]
            if isinstance(a, str) and isinstance(b, str):
                content[i] = a + b
                del content[i + 1]
            else:
                i += 1

        result = html.span(*content)
        if style:
            result.style = style
        return result

    def rewrite_spans(parent):
        # Figure out where to start rewriting. If Word numbering inserted a
        # span.marker, skip it.
        rewritable_content_start = 0
        if parent.content and not isinstance(parent.content[rewritable_content_start], str):
            first = parent.content[rewritable_content_start]
            if is_marker(first):
                rewritable_content_start += 1

        cls = parent.attrs['class']
        inherited_style = docx.styles[cls].full_style

        # Determine the style of each run of content in the paragraph.
        items = []
        for kid in parent.content[rewritable_content_start:]:
            if not isinstance(kid, str) and kid.name == 'span':
                run_style = inherited_style.copy()
                run_style.update(kid.style)
                if 'class' in kid.attrs:
                    run_style.update(docx.styles[kid.attrs['class']].full_style)
                items.append((kid.content, run_style))
            else:
                items.append(([kid], inherited_style))

        # Drop trailing whitespace at end of paragraph.
        while items and all(isinstance(ht, str) and ht.isspace() for ht in items[-1][0]):
            del items[-1]

        # If the paragraph begins and ends in the same font, treat that font
        # as the paragraph's font, which we will drop.
        paragraph_style = inherited_style.copy()
        if paragraph_style.get('font-family') == 'monospace':
            del paragraph_style['font-family']
        if items:
            start_font = items[0][1].get('font-family')
            if start_font is not None and start_font != 'monospace':
                end_font = items[-1][1].get('font-family')
                if start_font == end_font:
                    paragraph_style['font-family'] = start_font

        # Build the ranges.
        all_content = []
        ranges = collections.defaultdict(dict)
        current_style = {}

        def set_current_style_to(style):
            here = len(all_content)
            for prop, (start, old_val) in list(current_style.items()):
                if style.get(prop, not old_val) != old_val:
                    # note end of earlier style
                    ranges[start, here][prop] = old_val
                    del current_style[prop]
            for prop, val in style.items():
                if prop not in current_style:
                    # note start of new style
                    current_style[prop] = here, val
                else:
                    assert current_style[prop][1] == val
            assert {k: v for k, (_, v) in current_style.items()} == style

        for content, run_style in items:
            set_current_style_to({p: v for p, v in run_style.items() if paragraph_style.get(p) != v})
            all_content += content
        set_current_style_to({})

        # Convert ranges to a list.
        ranges = [(start, stop, style) for (start, stop), style in ranges.items()]
        ranges.sort(key=lambda triple: (triple[0], -triple[1]))

        def build_result(ranges, i0, i1):
            result = []
            content_index = i0
            while ranges:
                start, stop, style = ranges[0]
                assert i0 <= start < stop <= i1
                assert content_index <= start
                result += all_content[content_index:start]  # add any plain content

                # split 'ranges' into two parts
                inner_ranges = []
                after_ranges = []
                for triple in ranges[1:]:
                    r0, r1, rs = triple
                    assert start <= r0 < r1 <= i1
                    if r1 <= stop:
                        inner_ranges.append(triple)
                    elif stop <= r0:
                        after_ranges.append(triple)
                    else:
                        # the gross case, hopefully rare
                        inner_ranges.append((r0, stop, rs))
                        after_ranges.append((stop, r1, rs))

                # recurse to build the child, add that to the result
                child_content = build_result(inner_ranges, start, stop)
                result.append(new_span(child_content, style))

                content_index = stop
                ranges = after_ranges

            result += all_content[content_index:i1]  # add any trailing plain content
            return result

        return [parent.with_content_slice(rewritable_content_start,
                                          len(parent.content),
                                          build_result(ranges, 0, len(all_content)))]

    return doc.replace('p', rewrite_spans)

tag_names = {
    'ANNEX': 'h1.l1',
    'a2': 'h1.l2',
    'a3': 'h1.l3',
    'a4': 'h1.l4',
    # Algorithm styles are handled via their list attributes
    'Alg2': None,
    'Alg3': None,
    'Alg4': None,
    'Alg40': None,
    'Alg41': None,
    'Algorithm': None,
    'bibliography': 'li.bibliography-entry',
    'BulletNotlast': 'li',
    'Caption': 'figcaption',
    'DateTitle': 'h1',
    'ECMAWorkgroup': 'h1.ECMAWorkgroup',
    'Example': '.Note',
    'Figuretitle': 'figcaption',
    'Heading1': 'h1.l1',
    'Heading2': 'h1.l2',
    'Heading3': 'h1.l3',
    'Heading4': 'h1.l4',
    'Heading5': 'h1.l5',
    'Introduction': 'h1',
    'ListBullet': 'li.ul',
    'M0': None,
    'M4': None,
    'M20': 'div.math-display',
    'MathDefinition4': 'div.display',
    'MathSpecialCase3': 'li',
    'Note': '.Note',
    'RefNorm': 'p.formal-reference',
    'StandardNumber': 'h1.StandardNumber',
    'StandardTitle': 'h1',
    'Syntax': 'h2',
    'SyntaxDefinition': 'div.rhs',
    'SyntaxDefinition2': 'div.rhs',
    'SyntaxRule': 'div.lhs',
    'SyntaxRule2': 'div.lhs',
    'Tabletitle': 'figcaption',
    'TermNum': 'h1',
    'Terms': 'p.Terms',
    'zzBiblio': 'h1',
    'zzSTDTitle': 'div.inner-title'
}

heading_styles = {k for k, v in tag_names.items()
                        if v == 'h1' or v == 'h2' or (v is not None and v.startswith('h1.'))}

def map_body(doc, f):
    head, body = doc.content
    return doc.with_content([head, f(body)])

@Fixup
def fixup_lists(doc, docx):
    """ Group numbered paragraphs into lists. """

    # A List represents either an element that is <ol>, <ul>, or <body>.
    #
    # parent: The List that contains this List, or None if this List is the
    # document body.
    #
    # left_margin: The x coordinate of the list marker, in points.
    #
    # content: This ol/ul/body element's .content list.
    #
    # numId, ilvl: The OOXML numbering info for this element, used only for
    # assertions. (List structure is recovered from indentation alone.)
    #
    # marker_type: 'bullet' or an integer indicating the level of nesting, so
    # that an outermost ol.proc has marker_type=0, the first nested ol.block
    # gets marker_type=1, and so on. Used only to assert that the markers
    # are correct.
    #
    List = collections.namedtuple('List', ['parent', 'left_margin', 'content',
                                           'numId', 'ilvl', 'marker_type'])

    def without_numbering_info(style):
        """ Return a dictionary just like style but without numbering entries. """
        s = None
        for k in ('-ooxml-numId', '-ooxml-ilvl'):
            if k in style:
                if s is None:
                    s = style.copy()
                del s[k]
        if s is None:
            return style
        return s

    def fix_body(body):
        # In a single left-to-right pass over the document, group paragraphs
        # into lists.  We start and end lists based on indentation alone.
        result = []

        # current is the current innermost List.
        current = List(parent=None, left_margin=-1e300, content=result,
                       numId=0, ilvl=None, marker_type=None)
        def append_non_list_item(e):
            if current.parent is None:
                # The enclosing element is the <body>. Just add this element to it.
                current.content.append(e)
            else:
                # The enclosing element is a list. It can contain only list
                # items, so put this paragraph or list in with the preceding
                # list item.
                assert(ht_name_is(current.content[-1], 'li'))
                current.content[-1].content.append(e)

        def open_list(p, numId, ilvl, margin):
            nonlocal current

            assert margin >= current.left_margin

            s = docx.get_list_style_at_level(numId, ilvl)
            is_bulleted_list = s is not None and s.numFmt == 'bullet'
            if is_bulleted_list:
                lst = html.ul()
                marker_type = 'bullet'
            else:
                if current.parent is None or current.marker_type == 'bullet':
                    cls = 'proc'
                    marker_type = 0
                elif (margin > (current.left_margin + 0.75 * 72)
                      and p.content
                      and is_marker(p.content[0])
                      and p.content[0].content == ['1.\t']):
                    # Very deeply nested list with decimal numbering.
                    cls = 'nested proc'
                    marker_type = 0
                else:
                    cls = 'block'
                    marker_type = current.marker_type + 1
                lst = html.ol(class_=cls)

            append_non_list_item(lst)
            current = List(parent=current, left_margin=margin, content=lst.content,
                           numId=numId, ilvl=ilvl, marker_type=marker_type)

        def close_list():
            nonlocal current
            current = current.parent
            assert current is not None

        for i, p in enumerate(body.content):
            # Get numbering info for this paragraph.
            numId = ilvl = None
            if p.style and '-ooxml-numId' in p.style and p.style['-ooxml-numId'] != '0':
                numId = int(p.style['-ooxml-numId'])
                ilvl = int(p.style.get('-ooxml-ilvl', '0'))

            # Determine the indentation depth.
            margin = 0.0
            if p.style and '-ooxml-indentation' in p.style:
                margin_str = p.style['-ooxml-indentation']
                if margin_str.endswith('pt'):
                    margin = float(margin_str[:-2])

                    # work around https://bugs.ecmascript.org/show_bug.cgi?id=1713
                    if margin < 0:
                        margin = 0
                else:
                    assert margin_str == '0' and margin == 0.0

            # Figure out if this paragraph is a list item.
            is_list_item = (numId is not None
                            and numId != 0
                            and p.attrs.get('class') not in heading_styles
                            and len(p.content) != 0
                            and is_marker(p.content[0])
                            and # work around https://bugs.ecmascript.org/show_bug.cgi?id=1715
                                p.content[0].content[0] != '8.4.6.2\t')

            # Close any more-indented active lists.
            #
            # Since -ooxml-indentation refers to the indentation of the numbering, not the text,
            # treat non-numbered paragraphs as being an additional 36pt (1/2 inch) to the left.
            # This is not meant to make sense.
            #
            effective_margin = margin
            if not is_list_item:
                effective_margin -= 36
            while current.left_margin > effective_margin:
                close_list()

            if not is_list_item:
                if p.style and '-ooxml-numId' in p.style:
                    p = p.with_(style=without_numbering_info(p.style))
                append_non_list_item(p)
            else:
                # If it is indented more than the previous paragraph, open a
                # new list.
                if margin > current.left_margin:
                    if numId == current.numId:
                        assert ilvl > current.ilvl
                    open_list(p, numId, ilvl, margin)

                assert margin == current.left_margin

                # Change the <p> to <li>, strip the marker and class=, and
                # strip the -ooxml-numId/ilvl style.  Add the result to the
                # current list.
                attrs = p.attrs
                if 'class' in attrs:
                    attrs = attrs.copy()
                    del attrs['class']
                li = p.with_(name='li',
                             content=p.content[1:],
                             attrs=attrs,
                             style=without_numbering_info(p.style))
                current.content.append(li)

                # Assert that the marker HTML will generate for this list item
                # is the same as the one that appears in the Word doc.
                marker_str = p.content[0].content[0]
                if current.marker_type == 'bullet':
                    # U+F0B7 is not a Unicode character, this is Word nonsense
                    assert marker_str == '\uf0b7\t'
                else:
                    depth = current.marker_type
                    i = len(current.content)
                    formatters = [str, int_to_lower_letter, int_to_lower_roman]
                    marker_formatter = formatters[depth % 3]
                    # Bizarrely, there is no dot after lowercase-letter markers at nesting depth 4.
                    html_marker_str = marker_formatter(i) + ('.' if depth != 4 else '') + '\t'
                    if html_marker_str != marker_str:
                        warn("Word marker is {!r}, HTML will show {!r}".format(marker_str, html_marker_str))
                        print(li)

        return [body.with_content(result)]

    def contains_paragraphs(e):
        return e.name in ('body', 'td')

    return doc.find_replace(contains_paragraphs, fix_body)

unrecognized_styles = collections.defaultdict(int)

def ht_concat(c1, c2):
    """ Concatenate two content lists. """
    assert isinstance(c1, list)
    assert isinstance(c2, list)
    if not c1:
        return c2
    elif not c2:
        return c1
    elif isinstance(c1[-1], str) and isinstance(c2[0], str):
        return c1[:-1] + [c1[-1] + c2[0]] + c2[1:]
    else:
        return c1 + c2

@Fixup
def fixup_paragraph_classes(doc, docx):
    annex_counters = [0, 0, 0, 0]

    def replace_tag_name(e):
        num = e.style and e.style.get('-ooxml-numId', '0') != '0'
        default_tag = 'li' if num else 'p'

        if 'class' not in e.attrs:
            return [e.with_(name=default_tag)]

        cls = e.attrs['class']

        if cls not in tag_names:
            unrecognized_styles[cls] += 1

        # work around https://bugs.ecmascript.org/show_bug.cgi?id=1715
        if e.content and is_marker(e.content[0]) and e.content[0].content[0] == '8.4.6.2\t':
            cls = 'Heading4'

        attrs = e.attrs.copy()
        del attrs['class']
        tag = tag_names.get(cls)
        if tag is None:
            tag = default_tag
        elif '.' in tag:
            tag, _, attrs['class'] = tag.partition('.')
            if tag == '':
                tag = default_tag
        return [e.with_(name=tag, attrs=attrs)]

    return doc.replace('p', replace_tag_name)

@Fixup
def fixup_remove_empty_headings(doc, docx):
    def is_empty(item):
        if isinstance(item, str):
            return item.strip() == ''
        elif is_marker(item):
            # TODO - strip this special case out; I think the generated marker
            # will be empty in the one case where this matters.
            return True
        else:
            return all(is_empty(c) for c in item.content)

    def remove_if_empty(heading):
        if is_empty(heading):
            return []
        else:
            return [heading]

    return doc.replace('h1', remove_if_empty)

@contextlib.contextmanager
def marker_temporarily_removed(e):
    # Remove the marker, if any.
    marker = None
    if e.content and is_marker(e.content[0]):
        marker = e.content.pop(0)

    yield

    # Put the marker back, if any.
    if marker:
        e.content.insert(0, marker)

@InPlaceFixup
def fixup_element_spacing(doc, docx):
    """
    Change "A<i> B</i>" to "A <i>B</i>".

    That is, move all start tags to the right of any adjacent whitespace,
    and move all end tags to the left of any adjacent whitespace.

    The exceptions are <pre> and <span class="marker">. These elements are left alone.
    """

    def rebuild(parent):
        result = []
        def addstr(s):
            if result and isinstance(result[-1], str):
                result[-1] += s
            else:
                result.append(s)

        for k in parent.content:
            if isinstance(k, str):
                addstr(k)
            elif k.name == 'pre' or is_marker(k):
                # Don't mess with spaces in a pre element or a marker.
                result.append(k)
            else:
                with marker_temporarily_removed(k):
                    discard_space = k.is_block()
                    if k.content:
                        a = k.content[0]
                        if isinstance(a, str) and a[:1].isspace():
                            k.content[0] = a_text = a.lstrip()
                            if not discard_space:
                                addstr(a[:len(a) - len(a_text)])
                            if a_text == '':
                                del k.content[0]
                    if k.content or k.attrs or k.name not in {'span', 'i', 'b', 'sub', 'sup'}:
                        result.append(k)
                    if k.content:
                        b = k.content[-1]
                        if isinstance(b, str) and b[-1:].isspace():
                            k.content[-1] = b_text = b.rstrip()
                            if not discard_space:
                                addstr(b[len(b_text):])

        parent.content[:] = result

    def walk(e):
        with marker_temporarily_removed(e):
            for i, kid in e.kids():
                walk(kid)
            rebuild(e)

    walk(doc)


def doc_body(doc):
    body = doc.content[1]
    assert body.name == 'body'
    return body

def ht_name_is(ht, name):
    return not isinstance(ht, str) and ht.name == name

@InPlaceFixup
def fixup_sec_4_3(doc, docx):
    for parent, i, kid in all_parent_index_child_triples(doc):
        # Hack: Sections 4.3.{7,8,16} are messed up in the document. Wrong style. Fix it.
        if (kid.name == "h1"
            and i > 0
            and ht_name_is(parent.content[i - 1], 'h1')
            and (kid.content == ["built-in object"]
                 or kid.content == ["String value"]
                 or kid.content == ["standard object"])):
            kid.name = "p"
            kid.attrs['class'] = 'Terms'

        if kid.name == 'p' and kid.attrs.get('class') == 'Terms' and i > 0 and parent.content[i - 1].name == 'h1':
            h1_content = parent.content[i - 1].content
            kid.content.insert(0, '\t')
            for item in kid.content:
                if isinstance(item, str) and h1_content and isinstance(h1_content[-1], str):
                    h1_content[-1] += item
                else:
                    h1_content.append(item)
            del parent.content[i]

@InPlaceFixup
def fixup_hr(doc, docx):
    """ Replace <p><hr></p> with <hr>.

    Word treats an explicit page break as occurring within a paragraph rather
    than between paragraphs, and this leads to goofy markup which has to be
    fixed up.

    Precedes fixup_sections and fixup_strip_toc, which depend on the <hr> tags.
    """

    body = doc_body(doc)
    for i, a in body.kids('p'):
        for j, b in a.kids('hr'):
            rest = a.content[j + 1:]
            if all(isinstance(ht, str) and ht.isspace() for ht in rest):
                result = []
                if not all(isinstance(ht, str) and ht.isspace() for ht in a.content[:j]):
                    del a.content[j:]
                    result.append(a)
                result.append(b)
                body.content[i:i + 1] = result
                break

@InPlaceFixup
def fixup_intl_remove_junk(doc, docx):
    """ Remove doc ids that only matter within Ecma, and some junk that Word inserts """
    for parent, i, child in all_parent_index_child_triples(doc):
        if child.name == 'h1' and child.attrs.get('class') == 'ECMAWorkgroup':
            del parent.content[i]
        if child.name == 'h1' and child.attrs.get('class') == 'StandardNumber' and child.content[0] == 'ECMA-XXX':
            del parent.content[i]
        if child.name == 'div' and child.attrs.get('class') == 'inner-title':
            t = child.content[0].partition('INTERNATIONAL STANDARD\N{COPYRIGHT SIGN}\N{NO-BREAK SPACE}ISO/IEC')
            child.content[0], _, _ = t

@InPlaceFixup
def fixup_sections(doc, docx):
    """ Group h1 elements and subsequent elements of all kinds together into sections. """

    body_elt = doc_body(doc)
    body = body_elt.content

    level_re = re.compile(r'l[1-5]')

    def starts_with_section_number(s):
        return re.match(r'[1-9]|[A-Z]\.[1-9][0-9]*', s) is not None

    def heading_info(h):
        """
        h is an h1 element. Return a pair (sec_num, title).
        sec_num is the section number, as a string, or None.
        title is the title, another string, or None.
        """

        c = h.content
        if len(c) == 0:
            return None, None

        # If this heading has a marker, convert it to not be a marker.
        if is_marker(h.content[0]):
            c = h.content = ht_concat(h.content[0].content, h.content[1:])

        s = c[0]
        if not isinstance(s, str):
            return None, None
        s = s.lstrip()

        # Very special hack: The heading for section 9.1.1 contains a tab
        # character. I don't see a tab in the PDF, though. Snip it out.
        if s.startswith('9.1\t.1\t'):
            s = s.replace('9.1\t.1\t', '9.1.1\t')

        # Very special hack: Fix two typos in section numbers.
        if s.startswith('8.4,4') or s.startswith('15,13'):
            # Replace only this first comma with a dot.
            before, comma, after = s.partition(',')
            s = before + '.' + after

        num, tab, title = s.partition('\t')
        if tab == "":
            if 1 < len(c) and ht_name_is(c[1], "span") and c[1].attrs.get("class") == "section-status":
                return s.strip(), ''
            elif 1 < len(c) and ht_name_is(c[1], 'br'):
                # Special hack: Annex headings have line breaks inside the header.
                # Parse the heading and modify it in-place to be less horrible. :-P
                assert num.startswith('Annex')
                status = c[2]
                assert status in ('(informative)', '(normative)')
                assert ht_name_is(c[3], 'br')
                title = ''
                for item in c[4:]:
                    if ht_name_is(item, 'br'):
                        title += ' '
                    else:
                        assert isinstance(item, str)
                        title += item
                h.content = [num + '\t', html.span(status, class_="section-status"), " " + title.strip()]
                return num.strip(), title
            elif starts_with_section_number(s):
                parts = s.split(None, 1)
                if len(parts) == 2:
                    return tuple(parts)
                else:
                    # Note this can happen if the section number is followed by an element.
                    return s.strip(), ''
            else:
                return None, s

        return num.strip(), title

    def contains(a, b):
        """ True if section `a` contains section `b` as a subsection.
        `a` and `b` are section numbers, which are strings; but some sections
        do not have numbers, so either or both may be None.
        """
        return a is not None and (b is None or b.startswith(a + "."))

    def sec_num_to_id(num):
        if num.startswith('Annex\N{NO-BREAK SPACE}'):
            return num[6:]
        else:
            return num

    def wrap(sec_num, sec_title, start):
        """ Wrap the section starting at body[start] in a section element. """
        sec_id = sec_num_to_id(sec_num).strip() if sec_num else None

        j = start + 1
        while j < len(body):
            kid = body[j]
            if not isinstance(kid, str):
                if kid.name == 'div':
                    if (kid.attrs.get('id') == 'ecma-disclaimer'
                        or kid.attrs.get('class') == 'inner-title'):
                        # Don't let the introduction section eat up these elements.
                        break
                elif kid.name == 'hr':
                    # Don't let the copyright notice eat the table of contents.
                    if sec_title == "Copyright notice":
                        break
                elif kid.name == 'h1':
                    kid_num, kid_title = heading_info(kid)

                    # Hack: most numberless sections are subsections, but the
                    # Bibliography is not contained in any other section.
                    if kid_title != 'Bibliography' and contains(sec_id, kid_num):
                        # kid starts a subsection. Wrap it!
                        wrap(kid_num, kid_title, j)
                    else:
                        # kid starts the next section. Done!
                        break
            j += 1
        stop = j

        attrs = {}
        if sec_num is not None:
            attrs['id'] = "sec-" + sec_id
            span = html.span(
                html.a(sec_num, href="#sec-" + sec_id, title="link to this section"),
                class_="secnum")
            c = body[start].content
            idx = 0
            if c and is_marker(c[0]):
                idx = 1
            c[idx:idx + 1] = [span, ' ' + sec_title]

        # Actually do the wrapping.
        body[start:stop] = [html.section(*body[start:stop], **attrs)]

    for i, kid in body_elt.kids("h1"):
        num, title = heading_info(kid)
        wrap(num, title, i)

    # remove some h1 attributes that we don't need anymore (or never needed)
    for h in findall(doc, 'h1'):
        if h.style.get('-ooxml-numId') != None:
            del h.style['-ooxml-numId']
            del h.style['-ooxml-ilvl']
        if h.attrs.get('class') != None:
            del h.attrs['class']

@InPlaceFixup
def fixup_strip_toc(doc, docx):
    """ Delete the table of contents in the document.

    Leave an empty section tag which can be populated with auto-generated
    contents later.

    This must follow fixup_hr.

    Precedes fixup_generate_toc.
    """
    body = doc_body(doc)
    toc = html.section(id='contents')

    hr_iterator = body.kids("hr")
    i0, first_hr = next(hr_iterator)
    if spec_is_lang(docx):
        if ht_text(body.content[i0 + 1]).startswith('Copyright notice'):
            # Skip the copyright notice that appears at the front of ES5.1.
            # The table of contents is right after that.
            i0, first_hr = next(hr_iterator)
        i1, next_hr = next(hr_iterator)
        i1 += 1
    else:
        section_iterator = body.kids("section")
        i1, first_section = next(section_iterator)
    body.content[i0: i1] = [toc]

@InPlaceFixup
def fixup_tables(doc, docx):
    """ Turn highlighted td elements into th elements.

    Also, OOXML puts all table cell content in paragraphs; strip out the extra
    <p></p> tags.

    Precedes fixup_pre, which converts p elements containing only code into
    pre elements; we don't want code elements in tables to handled that way.
    """

    def is_negligible_css_property(name):
        return name == '-ooxml-indentation' or name.startswith('margin-')

    def is_vacuous(e):
        return not e.attrs and (e.style is None
                                or all(is_negligible_css_property(p) for p in e.style))

    for td in findall(doc, 'td'):
        if td.style and td.style.get('background-color') in ('#C0C0C0', '#D8D8D8'):
            td.name = 'th'
            del td.style['background-color']

        if len(td.content) == 1 and ht_name_is(td.content[0], 'p'):
            p = td.content[0]
            if p.style and p.style.get('background-color') in ('#C0C0C0', '#D8D8D8'):
                td.name = 'th'
                del p.style['background-color']
            if len(p.content) == 1 and ht_name_is(p.content[0], 'span'):
                span = p.content[0]
                if span.style and span.style.get('background-color') in ('#C0C0C0', '#D8D8D8'):
                    td.name = 'th'
                    del span.style['background-color']

            # If the p is vacuous, kill it.
            if is_vacuous(p):
                td.content = p.content

            # Ditto if it happens to contain an empty span.
            if len(td.content) == 1 and ht_name_is(td.content[0], 'span'):
                span = td.content[0]
                if td.name == 'th' and span.style:
                    # Delete redundant style info.
                    if span.style.get('font-family') == 'Times New Roman':
                        del span.style['font-family']
                    if span.style.get('font-weight') == 'bold':
                        del span.style['font-weight']
                    if span.style.get('font-style') == 'italic':
                        del span.style['font-style']
                if is_vacuous(span):
                    td.content = span.content

@Fixup
def fixup_table_formatting(doc, docx):
    """ Mark each table as either .real-table or .lightweight-table. """
    def fix_table(table):
        if table.attrs and 'class' in table.attrs:
            return [table]

        def has_borders_or_shading(e):
            s = e.style
            return s and any(k.startswith(('border', '-ooxml-border', 'background'))
                             for k in s)

        formatted = any(cell.name == 'th' or has_borders_or_shading(cell)
                          for _, row in table.kids('tr')
                            for _, cell in row.kids())
        attrs = table.attrs.copy()
        if formatted:
            attrs['class'] = 'real-table'
        else:
            attrs['class'] = 'lightweight-table'
        return [table.with_(attrs=attrs)]
    return doc.replace('table', fix_table)

@Fixup
def fixup_pre(doc, docx):
    """ Convert p elements containing only monospace font to pre.

    Precedes fixup_notes, which considers pre elements to be part of notes.
    """

    def is_code_para(e):
        if e.name == 'table':
            return None  # do not walk this subtree
        elif e.name == 'p' and len(e.content) == 1:
            [kid] = e.content
            return ht_name_is(kid, 'span') and kid.style and kid.style.get('font-family') == 'monospace'
        else:
            return False  # don't convert this element, but do walk the subtree

    def convert_para(e):
        return [e.with_(name='pre', content=e.content[0].content)]

    return doc.find_replace(is_code_para, convert_para)

@InPlaceFixup
def fixup_notes(doc, docx):
    """ Wrap each NOTE and EXAMPLE in div.note and wrap the labels "NOTE", "NOTE 2", etc. in span.nh. """

    def find_nh(p, strict=False):
        if len(p.content) == 0:
            if strict:
                warn("warning in fixup_notes: p.Note paragraph has no content")
            return None
        s = p.content[0]
        if not isinstance(s, str):
            if strict:
                warn("warning in fixup_notes: p.Note paragraph does not start with a string")
            return None
        else:
            left, tab, right = s.partition('\t')
            if tab == '':
                if strict:
                    warn('warning in fixup_notes: no tab in NOTE: ' + repr(s))
                return None
            elif not (left.startswith('NOTE') or left.startswith('EXAMPLE')):
                if strict:
                    warn('warning in fixup_notes: no "NOTE" or "EXAMPLE" in p.Note: ' + repr(s))
                return None
            else:
                return left, right

    def can_be_included(next_sibling):
        return (next_sibling.name in ('pre', 'ul')
                or (# total special case for the section on Number.prototype.toExponential()
                    next_sibling.name == 'ol'
                    and len(next_sibling.content) == 1)
                or (# the next sibling is p.Note but doesn't have a NOTE heading
                    next_sibling.name == 'p'
                    and next_sibling.attrs.get("class") == "Note"
                    and find_nh(next_sibling, strict=False) is None)
                or (# the next sibling is <p> and begins with a lowercase word
                    next_sibling.name == 'p'
                    and next_sibling.content
                    and isinstance(next_sibling.content[0], str)
                    and next_sibling.content[0].strip().split(None, 1)[0].islower()))

    for parent, i, p in all_parent_index_child_triples(doc):
        if p.name == 'p':
            # The Note class is unreliable: there are both false positives and
            # false negatives.  We only use it to emit warnings for the false
            # positives.
            has_note_class = p.attrs.get('class') == 'Note'
            nh_info = find_nh(p, strict=has_note_class)
            if nh_info:
                # This is a note! See if the word "NOTE", "NOTE 1", or "EXAMPLE" can be divided out into
                # a span.nh element. This should ordinarily be the case.
                nh, rest = nh_info
                assert p.content[0] == nh + '\t' + rest
                p.content[0] = ' ' + rest
                p.content.insert(0, html.span(nh, class_="nh"))

                # We don't need .Note anymore; remove it.
                if has_note_class:
                    del p.attrs['class']

                # Look for sibling elements that belong to the same note.
                j = i + 1
                while j < len(parent.content) and can_be_included(parent.content[j]):
                    if parent.content[j].attrs.get('class') == 'Note':
                        del parent.content[j].attrs['class']
                    j += 1

                # Wrap the whole note in a div.note element.
                parent.content[i:j] = [html.div(*parent.content[i:j], class_="note")]

def ht_text(ht):
    if isinstance(ht, str):
        return ht
    elif isinstance(ht, list):
        return ''.join(map(ht_text, ht))
    else:
        return ht_text(ht.content)

def find_section(doc, title):
    # super slow algorithm
    for sect in findall(doc, 'section'):
        if sect.content and ht_name_is(sect.content[0], 'h1'):
            h = sect.content[0]
            i = 0
            if i < len(h.content) and is_marker(h.content[i]):
                i += 1
            if i < len(h.content) and ht_name_is(h.content[i], 'span') and h.content[i].attrs.get('class') == 'secnum':
                i += 1
            s = ht_text(h.content[i:])
            if s.strip() == title:
                return sect
    raise ValueError("No section has the title " + repr(title))

@InPlaceFixup
def fixup_lang_15_10_2_2(doc, docx):
    """ Fix the ilvl attributes on the nested procedure in section 15.10.2.2
        of the Language specification.

    Precedes fixup_lists which consumes this data.
    """
    sect = find_section(doc, 'Pattern')
    assert len(sect.content) == 11
    assert [e.name for e in sect.content] == ['h1', 'p', 'li', 'li', 'li', 'li', 'li', 'li', 'li', 'li', 'div']
    for li in sect.content[4:10]:
        assert li.name == 'li'
        assert li.style['-ooxml-ilvl'] == '0'
        li.style['-ooxml-ilvl'] = '3'

def map_section(doc, title, fixup):
    hits = 0

    def match(e):
        nonlocal hits
        if e.name != 'section':
            if e.name in ('body', 'html'):
                return False  # no match, but visit children
            return None  # no match and skip entire subtree
        if len(e.content) == 0:
            return False
        h = e.content[0]
        if not ht_name_is(h, 'h1'):
            return False
        i = 0
        if i < len(h.content) and ht_name_is(h.content[i], 'span') and h.content[i].attrs.get('class') == 'marker':
            i += 1
        if i < len(h.content) and ht_name_is(h.content[i], 'span') and h.content[i].attrs.get('class') == 'secnum':
            i += 1
        s = ht_text(h.content[i:])
        if s.strip() == title:
            hits += 1
            return True
        else:
            return False

    def replacement(e):
        return [fixup(e)]

    result = doc.find_replace(match, replacement)
    if hits == 0:
        raise ValueError("could not find section to patch: {!r}".format(title))
    elif hits > 1:
        raise ValueError("map_section: found multiple sections with title {!r}".format(title))
    return result

@Fixup
def fixup_lang_15_12_3(doc, docx):
    """ Convert some paragraphs in section 15.12.3 of the Language specification into tables. """
    def is_target(e):
        return e.name == 'li' and any(p.name == 'p' and ht_text(p).startswith('backspace\t')
                                      for i, p in e.kids())

    def fix_target(li):
        def row(p):
            word, char = ht_text(p).strip().split('\t')
            return html.tr(html.td(word), html.td(html.span(char, class_="string value")))

        for i, p in li.kids():
            if p.name == 'p' and ht_text(p).startswith('backspace\t'):
                j = i + 1
                while j < len(li.content) and ht_name_is(li.content[j], 'p'):
                    j += 1
                tbl = html.table(*map(row, li.content[i:j]), class_='lightweight')
                return [li.with_content_slice(i, j, [tbl])]
        raise ValueError("fixup_lang_15_12_3: could not find text to patch in target list item")

    def fix_sect(sect):
        result = sect.find_replace(is_target, fix_target)
        if result is sect:
            raise ValueError("fixup_lang_15_12_3: could not find list item to patch in section")
        return result

    return map_section(doc, 'JSON.stringify ( value [ , replacer [ , space ] ] )', fix_sect)


def starts_with_marker(content):
    assert isinstance(content, list)
    return content and is_marker(content[0])

@InPlaceFixup
def fixup_list_paragraphs(doc, docx):
    """ Put some more space between list items in certain lists. """

    def is_block(ht):
        return not isinstance(ht, str) and ht.is_block()

    for parent, _, kid in all_parent_index_child_triples(doc):
        if kid.name == 'ul':
            ul = kid
            n = 0
            chars = 0
            for _, li in ul.kids('li'):
                # Check to see if this list item already contains a paragraph
                first = li.content[0]
                if not isinstance(first, str) and first.is_block():
                    chars = -1
                    break
                n += 1
                chars += len(ht_text(li))

            # If the average list item is any length, make paragraphs.
            if (spec_is_lang(docx) or parent.name != 'li') and chars / n > 80:
                for _, li in ul.kids('li'):
                    i = len(li.content)
                    while i > 0 and is_block(li.content[i - 1]):
                        i -= 1
                    li.content[:i] = [html.p(*li.content[:i])]

def replace_figure(doc, section_title, n, alt, width, height, has_svg=False):
    image = html.img(src="figure-{}.png".format(n), width=str(width), height=str(height), alt=alt)
    if has_svg:
        image = html.object(image, type="image/svg+xml", width=str(width), height=str(height),
                            data="figure-{}.svg".format(n))

    def f(sect):
        # Find the index of figure 1 within sect.content.
        c = sect.content
        for i, e in enumerate(c):
            if ht_name_is(e, 'figure') and i + 1 < len(c):
                caption = c[i + 1]
                if (ht_name_is(caption, 'figcaption')
                      and caption.content
                      and caption.content[0].startswith('Figure ' + str(n))):
                    # Found figure
                    figure = html.figure(image, caption)
                    return sect.with_content(c[:i] + [figure] + c[i + 2:])
        warn("figure {} not found".format(n))
        return sect

    return map_section(doc, section_title, f)

@Fixup
def fixup_figure_1(doc, docx):
    return replace_figure(doc, "Objects", 1,
                          alt="An image of lots of boxes and arrows.",
                          width=719, height=354,
                          has_svg=True)

@Fixup
def fixup_figure_2(doc, docx):
    return replace_figure(doc, "GeneratorFunction Objects", 2,
                          alt="A staggering variety of boxes and arrows.",
                          width=968, height=958)

@Fixup
def fixup_remove_picts(doc, docx):
    """ Remove div.w-pict elements. """

    def is_pict(e):
        return (e.name == 'div' and e.attrs.get('class') == 'w-pict')

    def rm_pict(e):
        # Remove the <div> element, but retain its contents.
        return e.content

    def is_pict_only_paragraph(e):
        return (e.name == 'p'
                and len(e.content) == 1
                and not isinstance(e.content[0], str)
                and is_pict(e.content[0]))

    def rm_pict_only_paragraph(e):
        # Remove the <p> and <div> elements, but retain their contents.
        return e.content[0].content

    return (doc
            .find_replace(is_pict_only_paragraph, rm_pict_only_paragraph)
            .find_replace(is_pict, rm_pict))

@InPlaceFixup
def fixup_figures(doc, docx):
    for parent, i, child in all_parent_index_child_triples(doc):
        if (child.name == 'figcaption'
              and i + 1 < len(parent.content)
              and ht_name_is(parent.content[i + 1], 'figure')):
            # add id to table captions that can have cross-references in word
            s = child.content[0]
            prefix = 'Table '
            if isinstance(s, str) and s.startswith(prefix):
                stop = len(prefix)
                while stop < len(s) and '0' <= s[stop] <= '9':
                    stop += 1
                table_id = s[len(prefix):stop]
                child.content[0] = html.span(id = 'table-' + table_id, * 'Table ' + table_id)
                rest = s[stop:]
                if rest:
                    child.content.insert(1, rest)
            # The iterator is actually ok with this mutation, but it's tricky.
            figure = parent.content[i + 1]
            del parent.content[i]
            figure.content.insert(0, child)

@InPlaceFixup
def fixup_remove_hr(doc, docx):
    """ Remove all remaining hr elements. """
    for parent, i, child in all_parent_index_child_triples(doc):
        if child.name == 'hr':
            del parent.content[i]

@InPlaceFixup
def fixup_title_page(doc, docx):
    """ Apply a fixup or two for the title page. """
    for parent, i, child in all_parent_index_child_triples(doc):
        if parent.name == 'p' and child.name == 'h1':
            # A p element shouldn't contain an h1, so make this an hgroup.
            parent.name = 'hgroup'
            parent.content = [k for k in parent.content if not ht_name_is(k, 'img')]
            if len(parent.content) != 6:
                continue

            # A few of the lines here are redundant.
            del parent.content[3:]

@Fixup
def fixup_lang_title_page_p_in_p(doc, docx):
    """
    Flatten a place where a few <p> elements appear inside another <p>
    (due to a <pict>).
    """
    def fix_p(p):
        if any(kid.is_block() for i, kid in p.kids()):
            last = None
            result = []
            for item in p.content:
                if not isinstance(item, str) and item.is_block():
                    last = None
                    result.append(item)
                else:
                    if last is None:
                        last = p.with_content([])
                        result.append(last)
                    last.content.append(item)
            return result
        else:
            return [p]
    return doc.replace('p', fix_p)

@InPlaceFixup
def fixup_html_head(doc, docx):
    head, body = doc.content
    assert ht_name_is(head, 'head')
    assert ht_name_is(body, 'body')
    head.content.insert(0, html.meta(charset='utf-8'))
    hgroup = next(findall(body, 'hgroup'))
    if spec_is_lang(docx):
        if '5.1' in ht_text(hgroup):
            title = "ECMAScript Language Specification - ECMA-262 Edition 5.1"
        else:
            title = "ECMAScript Language Specification ECMA-262 6th Edition - DRAFT"
    else:
        if version_is_intl_1_final(docx):
            title = "ECMAScript Internationalization API Specification - ECMA-402 Edition 1.0"
        else:
            title = "ECMAScript Internationalization API Specification - ECMA-402 Edition 1.0 - DRAFT"
    title = title.replace(' - ', ' \N{EN DASH} ')
    head.content.insert(1, html.title(title))
    if spec_is_lang(docx):
        stylesheet = 'es6-draft.css'
    else:
        stylesheet = 'es5.1.css'
    head.content.insert(2, html.link(rel='stylesheet', href=stylesheet))
    doc.attrs['lang'] = 'en-GB'

@InPlaceFixup
def fixup_lang_overview_biblio(doc, docx):
    sect = find_section(doc, "Overview")
    for i, p in enumerate(sect.content):
        if p.name == 'p':
            if p.content and p.content[0].startswith('Gosling'):
                break

    # First, strip the span element around the &trade; symbol.
    assert p.content[1].name == 'span'
    assert p.content[1].style == {'vertical-align': 'super'}
    assert p.content[1].content == ['\N{TRADE MARK SIGN}']
    s = p.content[0] + p.content[1].content[0] + p.content[2]

    # Make this paragraph a reference.
    p.attrs['class'] = 'formal-reference'

    # Italicize the title.
    people, dot_space, rest = s.partition('. ')
    assert dot_space
    title, dot_space, rest = rest.partition('. ')
    assert dot_space
    p.content[0:3] = [people + dot_space, html.span(title.strip(), class_="book-title"), dot_space + rest]

    # Fix up the second reference.
    i += 1
    p = sect.content[i]
    assert p.name == 'p'
    p.attrs['class'] = 'formal-reference'
    people, dot_space, rest = p.content[0].partition('. ')
    assert dot_space
    title, dot_space, rest = rest.partition('. ')
    assert dot_space
    p.content[0:1] = [people + dot_space, html.span(title.strip(), class_="book-title"), dot_space + rest]

    # Fix up the third reference.
    i += 1
    p = sect.content[i]
    assert p.name == 'p'
    assert p.content[0].startswith('IEEE')
    p.attrs['class'] = 'formal-reference'
    title, dot_space, rest = p.content[0].partition('.')
    assert dot_space
    p.content[0:1] = [html.span(title.strip(), class_="book-title"), dot_space + rest]

@Fixup
def fixup_simplify_formatting(doc, docx):
    """ Convert formatting spans into HTML markup that does the same thing.

    (Sometimes this converts to semantic markup, like using a var or span.nt
    element instead of an <i>; but it's just a shot in the dark.)

    This precedes fixup_lang_grammar_pre which looks for sub and span.nt elements.
    """

    nt_re = re.compile(r'\s+|\S+')

    def simplify_style_span(span):
        if span.attrs:
            return [span]
        if not span.style:
            return span.content

        style = span.style
        content = span.content

        if content == ["opt"] and style == {'font-family': 'sans-serif', 'vertical-align': 'sub'}:
            return [html.sub("opt")]
        elif style == {'font-family': 'monospace', 'font-weight': 'bold'}:
            return [html.code(*content)]
        elif (style == {'font-family': 'Times New Roman', 'font-style': 'italic'}
              and len(content) == 1
              and isinstance(content[0], str)):
            words = content[0].strip().split()
            if all(looks_like_nonterminal(w) for w in words):
                # Don't use words, because it's stripped.
                arr = []
                for s in nt_re.findall(content[0]):
                    if s.isspace():
                        arr.append(s)
                    else:
                        arr.append(html.span(s, class_="nt"))
                return arr
            else:
                return [html.var(*content)]
        elif style == {'font-family': 'Times New Roman', 'font-weight': 'bold'}:
            return [html.span(*content, class_='value')]

        style = style.copy()
        if style.get('font-style') == 'italic':
            content = [html.i(*content)]
            del style['font-style']
        if style.get('font-weight') == 'bold':
            content = [html.b(*content)]
            del style['font-weight']
        if style.get('vertical-align') == 'super':
            content = [html.sup(*content)]
            del style['vertical-align']
        if style.get('vertical-align') == 'sub':
            content = [html.sub(*content)]
            del style['vertical-align']
        if style:
            content = [html.span(*content)]
            content[0].style = style
        return content

    return doc.replace('span', simplify_style_span)

@InPlaceFixup
def fixup_lang_grammar_pre(doc, docx):
    """ Convert runs of div.lhs and div.rhs elements in doc to pre elements.

    Keep the text; throw everything else away.
    """

    def is_indented(p):
        if p.style:
            ind = p.style.get('-ooxml-indentation')
            if ind is not None and ind.endswith('pt'):
                ind = float(ind[:-2])
                return ind >= 18
        return False

    def is_grammar_block(p, continued):
        if ht_name_is(p, 'div'):
            return p.attrs.get('class') in ('lhs', 'rhs')
        elif ht_name_is(p, 'p'):
            # Some plain old paragraphs are also grammar.
            if len(p.content) == 0:
                return False
            elif is_nonterminal(p.content[0]) and find_inline_production_stop_index(p, 0) == len(p.content):
                return True
            elif continued and is_indented(p) and find_inline_grammar_stop_index(p, 0) == len(p.content):
                return True
            else:
                return False
        else:
            return False

    def lines(div):
        line = ''
        all_lines = []
        def visit(content):
            nonlocal line
            for ht in content:
                if isinstance(ht, str):
                    line += ht
                elif ht.name == 'br':
                    all_lines.append(line)
                    line = ''
                elif ht.name == 'sub' and ht.content == ['opt']:
                    line += '_opt'
                else:
                    visit(ht.content)

        visit(div.content)
        if line:
            all_lines.append(line)
        return all_lines

    def is_lhs(text):
        text = re.sub(r'( one of-?)?( See ((\d+|[A-Z])(.\d+)*|clause \d+))?$', '', text)
        if text.endswith(' one of'):
            text = text[:-7]
        return text.endswith(':')

    def strip_grammar_block(parent, i):
        j = i + 1
        while j < len(parent.content) and is_grammar_block(parent.content[j], True):
            j += 1
        syntax = ''
        for e in parent.content[i:j]:
            for line in lines(e):
                line = line.strip()
                line = ' '.join(line.split())

                if line.startswith("any "):
                    line = '[desc ' + line + ']'
                elif 'U+0000 through' in line:
                    k = line.index('U+0000 through')
                    line = line[:k] + '[desc ' + line[k:] + ']'

                if is_lhs(line):
                    syntax += '\n'  # blank line before
                else:
                    syntax += '    '  # indent each rhs
                syntax += line + '\n'

        # One-line productions
        if syntax.count('\n') == 1 and " :" in syntax:
            syntax = syntax.lstrip()

        # Hack - not all the paragraphs marked as syntax are actually
        # things we want to replace. So as a heuristic, only make the
        # change if the first line satisfied is_lhs.
        if syntax.startswith('    '):
            return

        parent.content[i:j] = [html.pre(syntax, class_="syntax")]

    def is_nonterminal(ht):
        if isinstance(ht, str):
            return False
        elif ht.name == 'i' or (ht.style and ht.style.get('font-style') == 'italic'):
            return looks_like_nonterminal(ht_text(ht).strip())
        elif ht.name == 'span':
            return ht.attrs.get('class') == 'nt' or (len(ht.content) == 1 and is_nonterminal(ht.content[0]))
        else:
            return False

    notin = "\N{NOT AN ELEMENT OF}"
    inline_grammar_re = re.compile(
        r'^\s*(?:$|\[empty\]|\[no\s*$|here\]|\[Lexical goal|\[lookahead ' + notin + r'|{|}|\])')

    def is_grammar_inline_at(parent, i):
        ht = parent.content[i]
        if isinstance(ht, str):
            m = inline_grammar_re.match(ht)
            if m is None:
                return False
            end = m.end()
            if end != len(ht):
                # Need to split, ew, mutation
                parent.content[i: i + 1] = [ht[:end], ht[end:]]
            return True
        elif ht.name == 'span':
            return ht.attrs.get('class') == 'nt' or (len(ht.content) == 1 and is_grammar_inline_at(ht, 0))
        elif ht.name == 'sub':
            return ht.content == ['opt']
        else:
            return ht.name in ('code', 'i', 'b')

    def inline_grammar_text(content):
        s = ''
        for ht in content:
            if isinstance(ht, str):
                s += ht
            elif ht.name == 'sub' and ht.content == ['opt']:
                s += '_opt'
            else:
                s += inline_grammar_text(ht.content)
        return s

    def find_inline_production_stop_index(e, i):
        # This algorithm is ugly. The only thing it has going for it is the
        # lack of evidence that something smarter would do a better job.

        # Skip any whitespace immediately following e.content[i]. If that
        # puts us at the end of e, there is no grammar production here;
        # return without doing anything.
        content = e.content
        j = i + 1
        if j >= len(content):
            return None
        if isinstance(content[j], str):
            if content[j].strip() == '':
                j += 1
                if j >= len(content):
                    return None

        # Hack: some productions are written (roughly) <b>::</b><code>.</code>
        # with no space between. Insert a space to make it work.
        free_pass = False
        eq = content[j]
        if (not isinstance(eq, str)
              and len(eq.content) > 1
              and isinstance(eq.content[0], str)
              and eq.content[0].startswith(':')
              and eq.content[0].rstrip(':') == ''):
            eq.content[0] += ' '
            free_pass = True

        # If we got a free pass, don't bother sanity-checking content[j].
        if not free_pass:
            jtext = ht_text(content[j]).lstrip()
            if not jtext.startswith(':') or jtext.split(None, 1)[0].rstrip(':') != '':
                return None

        return find_inline_grammar_stop_index(e, j + 1)

    def find_inline_grammar_stop_index(e, j):
        while j < len(e.content) and is_grammar_inline_at(e, j):
            j += 1
        return j

    def strip_grammar_inline(parent, i):
        """ Find a grammar production in parent, starting at parent.content[i].

        Replace it with a span.prod element.
        """

        j = find_inline_production_stop_index(parent, i)
        if j is None:
            return

        # Strip out all formatting and replace parent.content[i:j] with a new span.prod.
        content = parent.content
        text = inline_grammar_text(content[i:j])
        text = ' '.join(text.strip().split())
        content[i:j] = [html.span(text, class_='prod')]

        # Insert a space after, unless there already is one or we're at the end
        # of a paragraph.
        if i + 1 < len(content):
            next_ht = content[i + 1]
            if isinstance(next_ht, str):
                if not next_ht[:1].isspace():
                    content[i + 1] = ' ' + next_ht
            else:
                if next_ht.name != 'br' and not ht_text(next_ht)[:1].isspace():
                    content.insert(i + 1, ' ')

    for parent, i, child in all_parent_index_child_triples(doc):
        if is_grammar_block(child, False):
            strip_grammar_block(parent, i)
        elif is_nonterminal(child):
            strip_grammar_inline(parent, i)

@InPlaceFixup
def fixup_lang_grammar_post(doc, docx):
    """ Generate nice markup from the stripped-down pre.syntax elements
    created by fixup_lang_grammar_pre. """

    syntax_token_re = re.compile(r'''(?x)
        ( See \  (?:clause\ )? [0-9A-Z\.]+  # cross-reference
        | ((?:[A-Z]+[a-z]|uri)[A-Za-z]* (?:_opt)?)  # nonterminal
        | one\ of
        | but\ not\ one\ of
        | but\ not
        | or
        | \[no\ LineTerminator\ here\]
        | \[desc \  [^]]* \]
        | \[empty\]
        | \[Lexical\ goal\ [A-Z][A-Za-z]*\]
        | \[lookahead \  . [^]]* \]     # the . stands for &notin;
        | <[A-Z]+>                      # special character
        | [()]                          # unstick a parenthesis from the following token
        | ;\ _opt                       # a terminal is optional in just one case
        | [^ ]*                         # any other token
        )\s*
        ''')

    def markup_syntax(text, cls, xrefs=None):
        xref = None
        markup = []

        make_geq = cls in ('lhs', 'prod')

        i = 0
        while i < len(text):
            m = syntax_token_re.match(text, i)
            if markup:
                markup.append(' ')

            token = m.group(1)
            opt = token.endswith('_opt')
            if opt:
                token = token[:-4]

            if token.startswith('See '):
                assert xref is None
                xref = html.div(token, class_='gsumxref')
            elif m.group(2) is not None:
                # nonterminal
                markup.append(html.span(token, class_='nt'))
            elif token in ('one of', 'but not', 'but not one of', 'or'):
                markup.append(html.span(token, class_='grhsmod'))
            elif token == '[empty]':
                markup.append(html.span(token, class_='grhsannot'))
            elif token == '[no LineTerminator here]':
                markup.append(html.span('[no ',
                                        html.span('LineTerminator', class_='nt'),
                                        ' here]',
                                        class_='grhsannot'))
            elif token.startswith('[desc '):
                markup.append(html.span(token[6:-1].strip(), class_='gprose'))
            elif token.startswith('[Lexical goal '):
                assert token.endswith(']')
                n = len('[Lexical goal ')
                markup.append(html.span(token[:n],
                                        html.span(token[n:-1].strip(), class_='nt'),
                                        ']',
                                        class_='grhsannot'))
            elif token.startswith('[lookahead '):
                start = '[lookahead \N{NOT AN ELEMENT OF} '
                assert token.startswith(start)
                assert token.endswith(']')
                lookset = token[len(start):-1].strip()
                if lookset.isalpha() and lookset[0].isupper():
                    parts = [start, html.span(lookset, class_='nt'), ']']
                elif lookset[0] == '{' and lookset[-1] == '}':
                    parts = [start + '{']
                    for minitoken in lookset[1:-1].split(','):
                        if len(parts) > 1:
                            parts.append(', ')
                        parts.append(html.code(minitoken.strip(), class_='t'))
                    parts.append('}]')
                else:
                    parts = [token]
                markup.append(html.span(*parts, class_='grhsannot'))
            elif make_geq and token and token.rstrip(':') == '':
                markup.append(html.span(token, class_='geq'))
                make_geq = False
            elif token.startswith('<') and token.endswith('>'):
                if markup:
                    markup[-1] += token
                else:
                    markup.append(token)
            else:
                # A terminal.
                assert token
                markup.append(html.code(token, class_='t'))

            if opt:
                markup.append(html.sub('opt'))

            i = m.end()

        results = []
        if xref:
            results.append(xref)
        results.append(html.div(*markup, class_=cls))
        return results

    for parent, i, child in all_parent_index_child_triples(doc):
        if child.name == 'pre' and child.attrs.get('class') == 'syntax':
            divs = []
            [syntax] = child.content
            syntax = syntax.lstrip('\n')
            for production in syntax.split('\n\n'):
                lines = production.splitlines()
                assert not lines[0][:1].isspace()

                done = False
                if len(lines) == 1:
                    qq = markup_syntax(lines[0].strip(), 'prod')
                    if len(qq) == 1:
                        d = qq[0] if qq else html.div()
                        d.attrs['class'] = 'gp prod'
                        divs.append(d)
                        done = True

                if not done:
                    lines_out = markup_syntax(lines[0], 'lhs')
                    for line in lines[1:]:
                        assert line.startswith('    ')
                        lines_out += markup_syntax(line.strip(), 'rhs')
                    divs.append(html.div(*lines_out, class_='gp'))
            parent.content[i:i + 1] = divs
        elif child.name == 'span' and child.attrs.get('class') == 'prod':
            [syntax] = child.content
            [result] = markup_syntax(syntax.strip(), 'prod')
            child.content = result.content

@Fixup
def fixup_remove_margin_style(doc, docx):
    def is_margin_property(name):
        return name.startswith('margin-') or name in ('text-indent', '-ooxml-indentation')
    def has_margins(e):
        return e.style is not None and any(is_margin_property(k) for k in e.style)
    def without_margins(e):
        style = {k: v for k, v in e.style.items() if not is_margin_property(k)}
        return [e.with_(style=style)]
    return doc.find_replace(has_margins, without_margins)

@InPlaceFixup
def fixup_intl_insert_ids(doc, docx):
    """ Internationalization spec only: Create ids for the definitions of
        abstract operations that don't have their own sections, so that we
        can link to them directly.
    """
    ids_to_insert = [
        "CompareStrings",
        "FormatNumber", "ToRawPrecision", "ToRawFixed",
        "ToDateTimeOptions", "BasicFormatMatcher", "BestFitFormatMatcher",
        "FormatDateTime", "ToLocalTime"
    ]

    for p in findall(doc, 'p'):
        if len(p.content) > 0:
            content = p.content[0]
            prefix = 'When the '
            if isinstance(content, str) and content.startswith(prefix):
                for id in ids_to_insert:
                    if content.startswith(prefix + id + ' abstract operation'):
                        p.content.insert(0, prefix)
                        p.content.insert(1, html.span(id = id, *id))
                        p.content[2] = content[len(prefix + id):]

@InPlaceFixup
def fixup_links(doc, docx):
    algorithm_name_re = re.compile(r'''(?x)
        ^(
            %?[A-Z][A-Za-z0-9.%]{3,}
            (?: \s* \[ \s* @@[A-Za-z0-9.%]* \s* \]
              | \s* \[\[ \s* [A-Za-z0-9.%]* \s* \]\] )?
        )
        :?   # Ignore stray colon in a few headings
        (?:
            # Arguments (or something else in parentheses); or "Abstract Operation"; or both.
            (?: \s* \( .* \) )? \s* Abstract \s+ Operation \s*
            | \s* \( .* \)
        )
        $
    ''')

    def title_as_algorithm_name(title, sec_id):
        if sec_id.startswith('#sec-15.10.2.'):
            # This is ClassAtom or the name of some other nonterminal.
            # Not an algorithm or builtin-method name. Skip it for now.
            return None

        m = algorithm_name_re.match(title)
        if m is not None:
            return m.group(1)
        # Also allow matches like "ToPrimitive".
        if re.match(r'[A-Z][a-z]+[A-Z][A-Za-z0-9]+', title) is not None:
            return title

    algorithm_name_to_section = {}
    sections_by_title = {}
    for sect in findall(doc, 'section'):
        if 'id' in sect.attrs and sect.content and sect.content[0].name == 'h1':
            heading_content = sect.content[0].content[:]
            while (heading_content
                   and not isinstance(heading_content[0], str)
                   and heading_content[0].attrs['class'] in ('secnum', 'marker')):
                del heading_content[0]
            title = ht_text(heading_content).strip()
            title = " ".join(title.split())
            sec_id = '#' + sect.attrs['id']
            sections_by_title[title] = sec_id
            alg = title_as_algorithm_name(title, sec_id)
            if alg is not None:
                if any(pattern.format(alg) in sections_by_title
                       for pattern in ("{}",
                                       "{} Object Structure" ,
                                       "{} Constructors",
                                       "{} Objects",
                                       "The {} Constructor")):
                    # Kill this as an algorithm name; we shouldn't link it.
                    print("{}: superceded by previous section".format(alg))
                    algorithm_name_to_section[alg] = None
                elif alg in algorithm_name_to_section:
                    # Mark as a duplicate. (Don't delete the entry; that would
                    # be a bug if there are 3, 5, 7 sections with this name.)
                    print("{}: duplicate!".format(alg))
                    algorithm_name_to_section[alg] = None
                else:
                    print("{} => {}".format(alg, sec_id))
                    algorithm_name_to_section[alg] = sec_id

    fallback_section_titles = {
        "The List and Record Specification Type": "The List Specification Type",
        "The Completion Record Specification Type": "The Completion Specification Type",
        "Function Declaration Instantiation": None
    }

    # Normally, any section can link to any other section, even its own
    # subsection or parent section.  This dictionary overrides that.  Each
    # item (source, destination): False means that text in source does not
    # get linked to destination.
    linkability_overrides = {
        ('7.9.1', '7.9'): False,
        ('7.9.2', '7.9'): False
    }

    def can_link(source, target):
        s = source[5:] if source.startswith('#sec-') else source
        t = target[5:] if target.startswith('#sec-') else target
        return linkability_overrides.get((s, t), source != target)

    def has_word_breaks(s, i, text):
        # Check for word break before
        if i == 0:
            pass
        elif s[i-1].isalnum() or s[i-1] in '%@':
            return False

        # Check for word break after
        j = i + len(text)
        if j == len(s):
            pass
        elif text.endswith('('):
            pass
        elif s[j].isalnum():
            return False
        elif s[j:j + 2] == "]]":
            # don't treat the HasInstance in [[HasInstance]] as a separate word
            return False
        elif s[j:j + 1] == '.' and s[j + 1:j + 2].isalpha():
            # don't treat the foo in foo.bar as a separate word
            return False

        return True


    if version_is_5(docx):
        globalEnv = "The Global Environment"
    else:
        globalEnv = "Global Environment Records"

    specific_link_source_data_lang = [
        ("chain productions", "Context-Free Grammars"),
        ("chain production", "Context-Free Grammars"),

        # 5.2
        # Note that there's a hack below to avoid including the parenthesis in the <a> element.
        # We only want to match when the parenthesis is present, but it shouldn't be part of
        # the link.
        ("Assert", "Algorithm Conventions"),
        ("abs(", "Algorithm Conventions"),
        ("sign(", "Algorithm Conventions"),
        ("modulo", "Algorithm Conventions"),
        ("floor(", "Algorithm Conventions"),

        # clause 7
        ("automatic semicolon insertion (7.9)", "Automatic Semicolon Insertion"),
        ("automatic semicolon insertion (see 7.9)", "Automatic Semicolon Insertion"),
        ("automatic semicolon insertion", "Automatic Semicolon Insertion"),
        ("semicolon insertion (see 7.9)", "Automatic Semicolon Insertion"),

        # clause 8
        ("Type(", "Types"),
        ("List", "The List and Record Specification Type"),
        ("Completion Record", "The Completion Record Specification Type"),
        ("Completion", "The Completion Record Specification Type"),
        ("abrupt completion", "The Completion Record Specification Type"),
        ("NormalCompletion", "NormalCompletion"),
        ("ReturnIfAbrupt", "ReturnIfAbrupt"),
        ("Reference", "The Reference Specification Type"),
        ("GetBase", "The Reference Specification Type"),
        ("GetReferencedName", "The Reference Specification Type"),
        ("IsStrictReference", "The Reference Specification Type"),
        ("HasPrimitiveBase", "The Reference Specification Type"),
        ("IsPropertyReference", "The Reference Specification Type"),
        ("IsUnresolvableReference", "The Reference Specification Type"),
        ("unresolvable Reference", "The Reference Specification Type"),
        ("Unresolvable Reference", "The Reference Specification Type"),
        ("IsSuperReference", "The Reference Specification Type"),
        ("Property Descriptor", "The Property Descriptor Specification Type"),
        ("property key value", "The Object Type"),
        ("property key", "The Object Type"),
        ("internal data property", "Object Internal Methods and Internal Data Properties"),
        ("ordinary Function object", "Ordinary Function Objects"),
        ("ordinary function object", "Ordinary Function Objects"),
        ("Bound Function", "Bound Function Exotic Objects"),
        ("bound function", "Bound Function Exotic Objects"),
        ("[[BoundTargetFunction]]", "Bound Function Exotic Objects"),
        ("[[BoundThis]]", "Bound Function Exotic Objects"),
        ("[[BoundArguments]]", "Bound Function Exotic Objects"),
        ("Array exotic object", "Array Exotic Objects"),
        ("String exotic object", "String Exotic Objects"),
        ("Symbol exotic objects", "Symbol Exotic Objects"),
        ("exotic arguments object", "Exotic Arguments Objects"),

        # clause 9
        ("SameValue (according to 9.12)", "SameValue(x, y)"),
        ("the SameValue algorithm (9.12)", "SameValue(x, y)"),
        ("the SameValue Algorithm (9.12)", "SameValue(x, y)"),
        ("Get(", "Get (O, P)"),
        ("Put(", "Put (O, P, V, Throw)"),

        # 10.1
        ("strict mode code (see 10.1.1)", "Strict Mode Code"),
        ("strict mode code", "Strict Mode Code"),
        ("strict code", "Strict Mode Code"),
        ("base code", "Strict Mode Code"),

        # 10.2
        ("Lexical Environment", "Lexical Environments"),
        ("lexical environment", "Lexical Environments"),
        ("outer environment reference", "Lexical Environments"),
        ("outer lexical environment reference", "Lexical Environments"),
        ("environment record (10.2.1)", "Environment Records"),
        ("Environment Record", "Environment Records"),
        ("declarative environment record", "Environment Records"),
        ("Declarative Environment Record", "Environment Records"),
        ("Object Environment Record", "Environment Records"),
        ("object environment record", "Environment Records"),
        ("the global environment", globalEnv),
        ("the Global Environment", globalEnv),

        # 10.3
        ("LexicalEnvironment", "Execution Contexts"),
        ("VariableEnvironment", "Execution Contexts"),
        ("ThisBinding", "Execution Contexts"),
        ("Identifier Resolution as specified in 10.3.1", "Identifier Resolution"),
        ("Identifier Resolution(10.3.1)", "Identifier Resolution"),

        # 10.5
        ("Declaration Binding Instantiation", "Declaration Binding Instantiation"),
        ("declaration binding instantiation (10.5)", "Declaration Binding Instantiation"),
        ("Function Declaration Binding Instantiation", "Function Declaration Instantiation"),

        # clause 14
        ("Directive Prologue", "Directive Prologues and the Use Strict Directive"),
        ("Use Strict Directive", "Directive Prologues and the Use Strict Directive"),

        # clause 15
        ("direct call (see 15.1.2.1.1) to the eval function", "Direct Call to Eval"),

        # 15.3
        #("Function.prototype.bind", "Function.prototype.bind (thisArg [, arg1 [, arg2, \N{HORIZONTAL ELLIPSIS}]])"),

        # 15.9
        ("this time value", "Properties of the Date Prototype Object"),
        ("time value", "Time Values and Time Range"),
        ("Day(", "Day Number and Time within Day"),
        ("msPerDay", "Day Number and Time within Day"),
        ("TimeWithinDay", "Day Number and Time within Day"),
        ("DaysInYear", "Year Number"),
        ("TimeFromYear", "Year Number"),
        ("YearFromTime", "Year Number"),
        ("InLeapYear", "Year Number"),
        ("MonthFromTime", "Month Number"),
        ("DayWithinYear", "Month Number"),
        ("DateFromTime", "Date Number"),
        ("WeekDay", "Week Day"),
        ("LocalTZA", "Local Time Zone Adjustment"),
        ("DaylightSavingTA", "Daylight Saving Time Adjustment"),
        ("LocalTime", "Local Time"),
        ("UTC(", "Local Time"),
        ("HourFromTime", "Hours, Minutes, Second, and Milliseconds"),
        ("MinFromTime", "Hours, Minutes, Second, and Milliseconds"),
        ("SecFromTime", "Hours, Minutes, Second, and Milliseconds"),
        ("msFromTime", "Hours, Minutes, Second, and Milliseconds"),
        ("msPerSecond", "Hours, Minutes, Second, and Milliseconds"),
        ("msPerMinute", "Hours, Minutes, Second, and Milliseconds"),
        ("msPerHour", "Hours, Minutes, Second, and Milliseconds")
    ]

    specific_link_source_data_intl = [
        # clause 5
        ("List", "Notational Conventions"),
        ("Record", "Notational Conventions")
    ]

    non_section_ids_lang = {
    }

    non_section_ids_intl = {
        "CompareStrings": "CompareStrings",
        "FormatNumber": "FormatNumber",
        "ToRawPrecision": "ToRawPrecision",
        "ToRawFixed": "ToRawFixed",
        "ToDateTimeOptions": "ToDateTimeOptions",
        "BasicFormatMatcher": "BasicFormatMatcher",
        "BestFitFormatMatcher": "BestFitFormatMatcher",
        "FormatDateTime": "FormatDateTime",
        "ToLocalTime": "ToLocalTime",
        "15.9.1.8": "http://ecma-international.org/ecma-262/5.1/#sec-15.9.1.8",
        "introduction of clause 15": "http://ecma-international.org/ecma-262/5.1/#sec-15",
    }

    # Build specific_links from algorithm_name_to_section,
    # specific_link_source_data, sections_by_title, and
    # fallback_section_titles. This occurs in four easy steps.
    #
    # 1: Figure out which source data to use.
    if spec_is_lang(docx):
        specific_link_source_data = specific_link_source_data_lang
        non_section_ids = non_section_ids_lang
    else:
        specific_link_source_data = specific_link_source_data_intl
        non_section_ids = non_section_ids_intl

    # 2: Create some data structures.
    specific_link_dict = dict(specific_link_source_data)
    non_section_id_hrefs = []
    for id in non_section_ids:
        non_section_id_hrefs.append(non_section_ids[id])

    # 3. Build specific_links using the specific_link_source_data.
    specific_links = []
    for text, title in specific_link_source_data:
        if title in sections_by_title:
            sec = sections_by_title[title]
        elif spec_is_lang(docx):
            target_title = fallback_section_titles[title]
            if target_title is None:
                continue
            sec = sections_by_title[target_title]
        else:
            sec = non_section_ids_intl[title]
            if not sec.startswith('http'):
                sec = '#' + sec
        specific_links.append((text, sec))

    # 4. Add additional specific_links based on algorithm names that appear in
    # section headings.
    algorithm_pairs = sorted(algorithm_name_to_section.items(),
                             key=lambda pair: (-len(pair[0]), pair[0]))
    for alg, sec_id in algorithm_pairs:
        if sec_id is not None:
            if alg in specific_link_dict:
                warn("section {} {} is also the target of a specific_link entry ({!r}: {!r})".format(
                    sec_id, alg, alg, specific_link_dict[alg]))
            else:
                specific_links.append((alg, sec_id))

    # Assert that the specific_links above make sense; that is, that each link
    # with a "(7.9)" or "(see 7.9)" in it actually points to the named section.
    #
    # A warning here means sections were renumbered. Any number of things can
    # be wrong in the wake of such a change. :)
    #
    for text, target in specific_links:
        m = re.search(r'\((?:see )?([1-9][0-9]*(?:\.[1-9][0-9]*)*)\)', text)
        if m is not None:
            if target != '#sec-' + m.group(1):
                warn("text refers to section number " + repr(m.group(1)) + ", but actual section is " + repr(target))

    all_ids = set([kid.attrs['id'] for _, _, kid in all_parent_index_child_triples(doc) if 'id' in kid.attrs])

    SECTION = r'([1-9A-Z][0-9]*(?:\.[1-9][0-9]*)+|[Cc]lause\s+[1-9][0-9]*|[Aa]nnex\s+[A-Z])'
    def compile(re_source):
        return re.compile(re_source.replace("SECTION", SECTION))

    section_link_regexes_lang = list(map(compile, [
        # Match " (11.1.5)" and " (see 7.9)"
        # The space is to avoid matching "(3.5)" in "Math.round(3.5)".
        r' \(((?:see )?SECTION)\)',

        # Match "See 11.5" and "See clause 13" in span.gsumxref.
        r'^(See SECTION)$',

        # Match "Clause 8" in "as defined in Clause 8 of this specification"
        # and many other similar cases.
        r'(?:see|See|in|of|to|from|below|and) (SECTION)(?:$|\.$|[,:) ]|\.[^0-9])',

        # Match "(Clause 16)", "(see clause 6)".
        r'(?i)(?:)\((?:but )?((?:see\s+(?:also\s+)?)?clause\s+([1-9][0-9]*))\)',
        #r'(?i)(?:)\((?:but )?((?:see\s+(?:also\s+)?)?SECTION)\)',

        # Match the first section number in a parenthesized list "(13.3.5, 13.4, 13.6)"
        r'\((SECTION),\ ',

        # Match the first section number in a list at the beginning of a paragraph, "12.14:" or "12.7, 12.7:"
        r'^(SECTION)[,:]',

        # Match the second or subsequent section number in a parenthesized list.
        r', (SECTION)[,):]',

        # Match the penultimate section number in lists that don't use the
        # Oxford comma, like "13.3, 13.4 and 13.5"
        r' (SECTION) and\b',

        # Some cross-references are marked with Word fields.
        # In the Language spec, all REF fields are stripped out at this point
        # whether they refer to a section or not; hence the very strange and precarious
        # (SECTION|) in this regexp -- to allow group 2 to match the empty string.
        r'{ REF \w+ (?:\\r )?\\h }' + '\N{LEFT-TO-RIGHT MARK}' + r'?((SECTION|))',

        r'((Table [1-9][0-9]*))'
    ]))

    section_link_regexes_intl = list(map(compile, [
        # in the Internationalization spec, all internal cross references
        # are marked as such, so we get ugly but easy-to-find text after transformation
        r'\{ REF _Ref[0-9]+ \\r \\h \}(([0-9]+(\.[0-9]+)*))',
        r'\{ REF _Ref[0-9]+ \\h \}((Table [0-9]+))',
        # we need some external references as well
        r'((ES5,\s+[0-9]+(\.[0-9]+)*))',
        r'((15\.9\.1\.8))',
        r'((introduction of clause 15))',
    ]))

    # Disallow . ( ) , at the end since they usually aren't meant as part of the URL.
    url_re = re.compile(r'https?://[0-9A-Za-z;/?:@&=+$,_.!~*()\'-]+[0-9A-Za-z;/?:@&=+$_!~*\'-]')

    xref_re = re.compile(r'(.*)\{ REF _Ref[0-9]+ (\\r )?\\h \}' + '\N{LEFT-TO-RIGHT MARK}' + r'?')

    def find_link(s, current_section):
        best = None
        for text, target in specific_links:
            i = s.find(text)
            if (i != -1
                and can_link(current_section, target)  # don't link sections to themselves
                and has_word_breaks(s, i, text)
                and (best is None or i < best[0])):
                # New best hit.
                n = len(text)
                if text.endswith('('):
                    n -= 1
                best = i, i + n, target

        if spec_is_lang(docx):
            section_link_regexes = section_link_regexes_lang
        else:
            section_link_regexes = section_link_regexes_intl

        for link_re in section_link_regexes:
            m = link_re.search(s)
            while m is not None:
                link_text = m.group(2)
                if link_text in non_section_ids:
                    id = non_section_ids[link_text]
                elif link_text.startswith('ES5, '):
                    id = "http://ecma-international.org/ecma-262/5.1/#sec-" + link_text[5:]
                elif link_text.startswith('Table '):
                    id = 'table-' + link_text[6:]
                elif link_text == "":
                    id = None
                else:
                    # Get the target section id.
                    sec_num = link_text
                    if sec_num.lower().startswith('clause'):
                        sec_num = sec_num[6:].lstrip()
                    elif sec_num.lower().startswith('annex'):
                        sec_num = sec_num[5:].lstrip()
                    id = "sec-" + sec_num

                if id is not None and id not in all_ids and not id.startswith('http'):
                    warn("no such section: " + m.group(2))
                    m = link_re.search(s, m.end(1))
                else:
                    if id is not None and not id.startswith('http'):
                        id = "#" + id
                    hit = m.start(1), m.end(1), id
                    if best is None or hit < best:
                        best = hit
                    break

        m = url_re.search(s)
        if m is not None:
            hit = m.start(), m.end(), m.group(0)
            if best is None or hit < best:
                best = hit

        return best

    def linkify(parent, i, s, current_section):
        while s:
            m = find_link(s, current_section)
            if m is None:
                return
            start, stop, href = m
            if start > 0:
                prefix = s[:start]
                m = xref_re.match(prefix)
                if m:
                    prefix = m.group(1)
                parent.content.insert(i, prefix)
                i += 1

            if href is not None:
                assert (not href.startswith('#')
                        or href[1:] in all_ids
                        or href[1:] in non_section_id_hrefs)
                parent.content[i] = html.a(href=href, *s[start:stop])
                i += 1
            else:
                del parent.content[i]
            s = s[stop:]
            if s:
                parent.content.insert(i, s)

    def visit(e, current_section):
        id = e.attrs.get('id')
        if id is not None:
            current_section = '#' + id

        for i, kid in enumerate(e.content):
            if isinstance(kid, str):
                if current_section is not None:  # don't linkify front matter, etc.
                    linkify(e, i, kid, current_section)
            elif kid.name == 'a' and 'href' in kid.attrs:
                # Yo dawg. No links in links.
                pass
            elif kid.name == 'h1':
                # Don't linkify headings.
                pass
            else:
                visit(kid, current_section)


    visit(doc_body(doc), None)

@InPlaceFixup
def fixup_generate_toc(doc, docx):
    """ Generate a table of contents from the section headings. """

    def make_toc_list(e, depth=0):
        sublist = []
        for _, sect in e.kids("section"):
            if sect.attrs.get('id') != 'contents':
                sect_item = make_toc_for(sect, depth + 1)
                if sect_item:
                    sublist.append(html.li(*sect_item))
        if sublist:
            return [html.ol(*sublist, class_="toc")]
        else:
            return []

    def make_toc_for(sect, depth):
        if not sect.content:
            return []
        h1 = sect.content[0]
        if isinstance(h1, str) or h1.name != 'h1' or h1.content in (['Static Semantics'], ['Runtime Semantics']):
            return []

        output = []

        # Copy the content of the header.
        # TODO - make this clone enough to rip out the h1>span>a title= attribute
        # TODO - make a link when there isn't one
        output += h1.content[:]  # shallow copy, nodes may appear in tree multiple times

        # Find any subsections.
        if depth < 3:
            output += make_toc_list(sect, depth)

        return output

    body = doc_body(doc)
    for i, section in body.kids('section'):
        if section.attrs.get('id') == 'contents':
            break
    else:
        raise ValueError("Cannot find body>section#contents.")

    assert not section.content
    section.content = [html.h1("Contents")] + make_toc_list(doc_body(doc))

@Fixup
def fixup_add_disclaimer(doc, docx):
    div = html.div
    p = html.p
    strong = html.strong
    em = html.em
    i = html.i
    a = html.a

    if spec_is_lang(docx) and version_is_51_final(docx):
        disclaimer = div(
            p("This is the HTML rendering of ", i("ECMA-262 Edition 5.1, The ECMAScript Language Specification"), "."),
            p("The PDF rendering of this document is located at ",
              a("http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf",
                href="http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf"),
              "."),
            p("The PDF version is the definitive specification. Any discrepancies between this HTML version "
              "and the PDF version are unintentional."),
            id="unofficial")
        position = 1
    elif spec_is_lang(docx):
        disclaimer = div(
            p(strong("This is ", em("not"), " the official ECMAScript Language Specification.")),
            p("The most recent final ECMAScript standard is Edition 5.1, the PDF document located at ",
              a("http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf",
                href="http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf"),
              "."),
            p("This is a draft of the next edition of the standard."),
            p("This page is based on the current working draft published at ",
              a("http://wiki.ecmascript.org/doku.php?id=harmony:specification_drafts",
                href="http://wiki.ecmascript.org/doku.php?id=harmony:specification_drafts"),
              ". The program used to convert that Word doc to HTML is a custom-piled heap of hacks. "
              "It may have stripped out or garbled some of the formatting that makes "
              "the specification comprehensible. You can help improve the program ",
              a("here", href="https://github.com/jorendorff/es-spec-html"),
              "."),

            # (U+2019 is RIGHT SINGLE QUOTATION MARK, the character you're supposed to use for an apostrophe.)
            p("For copyright information, see Ecma International\u2019s legal disclaimer in the document itself."),
            id="unofficial")
        position = 0
    elif spec_is_intl(docx) and version_is_intl_1_final(docx):
        disclaimer = div(
            p("This is the HTML rendering of ", i("ECMA-402 Edition 1.0, ECMAScript Internationalization API Specification"), "."),
            p("The PDF rendering of this document is located at ",
              a("http://www.ecma-international.org/ecma-402/1.0/ECMA-402.pdf",
                href="http://www.ecma-international.org/ecma-402/1.0/ECMA-402.pdf"),
              "."),
            p("The PDF version is the definitive specification. Any discrepancies between this HTML version "
              "and the PDF version are unintentional."),
            id="unofficial")
        position = 1
    else:
        assert spec_is_intl(docx)
        disclaimer = div(
            p(strong("This is ", em("not"), " the official ECMAScript Internationalization API Specification.")),
            p("This is a draft of this standard."),
            p("This page is based on the current draft published at ",
              a("http://wiki.ecmascript.org/doku.php?id=globalization:specification_drafts",
                href="http://wiki.ecmascript.org/doku.php?id=globalization:specification_drafts"),
              ". The program used to convert that Word doc to HTML is a custom-piled heap of hacks. "
              "It may have stripped out or garbled some of the formatting that makes "
              "the specification comprehensible. You can help improve the program ",
              a("here", href="https://github.com/jorendorff/es-spec-html"),
              "."),
            p("For copyright information, see Ecma International\u2019s legal disclaimer in the document itself."),
            id="unofficial")
        position = 0

    def with_disclaimer(body):
        content = body.content[:position] + [disclaimer] + body.content[position:]
        return body.with_content(content)

    return map_body(doc, with_disclaimer)

@InPlaceFixup
def fixup_add_ecma_flavor(doc, docx):
    if version_is_intl_1_final(docx):
        hgroup = doc_body(doc).content[0]
        title = hgroup.content[0]
        del hgroup.content[0]
        hgroup.content[0].content.insert(0, "Standard ")
        hgroup.content.append(title)
        hgroup.content[0].style["color"] = "#ff6600"
        hgroup.content[1].style["color"] = "#ff6600"
        hgroup.content[2].style["color"] = "#ff6600"
        hgroup.content[2].style["font-size"] = "250%"
        hgroup.content[2].style["margin-top"] = "20px"
        doc_body(doc).content.insert(0, html.img(src="Ecma_RVB-003.jpg", alt="Ecma International Logo.",
            height="146", width="373"))


# === Main

def get_fixups(docx):
    yield fixup_strip_empty_paragraphs
    yield fixup_add_numbering
    yield fixup_list_styles
    yield fixup_formatting
    yield fixup_lists
    yield fixup_paragraph_classes
    yield fixup_remove_empty_headings
    yield fixup_element_spacing
    yield fixup_sec_4_3
    yield fixup_hr
    if spec_is_intl(docx):
        yield fixup_intl_remove_junk
    yield fixup_sections
    yield fixup_strip_toc
    yield fixup_tables
    yield fixup_table_formatting
    yield fixup_pre
    yield fixup_notes
    if spec_is_lang(docx):
        yield fixup_lang_15_12_3
    yield fixup_list_paragraphs
    if spec_is_lang(docx):
        yield fixup_figure_1
        yield fixup_figure_2
    yield fixup_remove_picts
    yield fixup_figures
    yield fixup_remove_hr
    yield fixup_title_page
    yield fixup_lang_title_page_p_in_p
    yield fixup_html_head
    if spec_is_lang(docx):
        yield fixup_lang_overview_biblio
    yield fixup_simplify_formatting
    if spec_is_lang(docx):
        yield fixup_lang_grammar_pre
        yield fixup_lang_grammar_post
    yield fixup_remove_margin_style
    if spec_is_intl(docx):
        yield fixup_intl_insert_ids
    yield fixup_links
    yield fixup_generate_toc
    yield fixup_add_disclaimer
    yield fixup_add_ecma_flavor

def fixup(docx, doc):
    logdir = "_fixup_log"
    if os.path.isdir(logdir):
        print("logging enabled")
    else:
        logdir = None

    for f in get_fixups(docx):
        print(f.name)
        t0 = time.time()
        doc = f(doc, docx)
        if logdir:
            filename = os.path.join(logdir, f.name + ".html")
            print("writing " + filename)
            html.save_html(filename, doc, strict=False)
        t1 = time.time()
        print("done ({} msec)".format(int(1000 * (t1 - t0))))
    return doc

