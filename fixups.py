import htmodel as html
import collections, re
from warnings import warn

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

def has_bullet(docx, p):
    if not p.style:
        return False
    numId = p.style.get('-ooxml-numId')
    if numId is None:
        return False
    ilvl = p.style.get('-ooxml-ilvl')
    s = docx.get_list_style_at(numId, ilvl)
    return s is not None and s.numFmt == 'bullet'

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

def fixup_formatting(doc, styles):
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

    run_style_properties = {
        'font-style',
        'font-weight',
        'font-family',
        'vertical-align'
    }

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

        if style == {'font-style': 'italic'}:
            return html.i(*content)
        elif style == {'font-weight': 'bold'}:
            return html.b(*content)
        elif style == {'vertical-align': 'super'}:
            return html.sup(*content)
        elif style == {'vertical-align': 'sub'}:
            return html.sub(*content)
        elif content == ["opt"] and style == {'font-family': 'sans-serif', 'vertical-align': 'sub'}:
            return html.sub("opt")
        elif style == {'font-family': 'monospace', 'font-weight': 'bold'}:
            return html.code(*content)
        else:
            result = html.span(*content)
            result.style = style
            return result

    def rewrite_spans(parent):
        spans = parent.content[:]  # copies the array

        # The role of inherited style is a little weird here.
        # We use it only to throw away pointless run markup.
        cls = parent.attrs.get('class', 'Normal')
        inherited_style = styles[cls].full_style

        if parent.style:
            # Delete w:rPr properties from the paragraph's style. As far as I
            # can tell they are always spurious; Word seems to ignore them.
            for prop, value in list(parent.style.items()):
                if prop in run_style_properties:
                    del parent.style[prop]

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

        # Build ranges.
        for kid in spans:
            if not isinstance(kid, str) and kid.name == 'span':
                run_style = kid.style
                if 'class' in kid.attrs:
                    run_style = run_style.copy()
                    run_style.update(styles[kid.attrs['class']].full_style)
                set_current_style_to({p: v for p, v in run_style.items() if inherited_style.get(p) != v})
                all_content += kid.content
            else:
                set_current_style_to({})
                all_content.append(kid)
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

        parent.content[:] = build_result(ranges, 0, len(all_content))

    for p in findall(doc, 'p'):
        # We assert the spans don't have attrs because we are going to
        # rewrite these guys retaining only the style. This fixup needs to
        # happen early enough in rewriting that this isn't a problem; it also
        # has to be early so that other markup doesn't get in the way.
        for i, kid in p.kids():
            if kid.name == 'span':
                assert len(kid.attrs) == 0 or list(kid.attrs.keys()) == ['class']
        rewrite_spans(p)

unrecognized_styles = collections.defaultdict(int)

def fixup_paragraph_classes(doc):
    annex_counters = [0, 0, 0, 0]
    def munge_annex_heading(e, cls):
        # Special case. Rather than implement OOXML numbering and Word
        # {SEQ} macros to the extent we'd need to generate the annex
        # headings, we fake it.
        
        # Figure out what level heading we are.
        if cls == 'ANNEX':
            level = 0
        else:
            level = int(cls[1:]) - 1
        assert level < len(annex_counters)

        # Bump the counter for this level; reset all the others to zero.
        annex_counters[level] += 1
        for i in range(level + 1, len(annex_counters)):
            annex_counters[i] = 0

        e.name = 'h1'
        letter = chr(ord('A') + annex_counters[0] - 1)
        if level == 0:
            # Parse the current content of the heading.
            i = 0
            content = e.content

            assert ht_name_is(content[i], 'br')
            i += 1

            status = content[i]
            status = re.sub(r'{SEQ .* }', '', status)
            assert status in ('(informative)', '(normative)')
            i += 1

            assert ht_name_is(content[i], 'br')
            i += 1

            while ht_name_is(content[i], 'br') or (isinstance(content[i], str)
                                                   and re.match(r'^{SEQ .* }$', content[i])):
                i += 1

            def ht_append(content, ht):
                if isinstance(ht, str) and content and isinstance(content[-1], str):
                    content[-1] += ht
                else:
                    content.append(ht)

            title = []
            while i < len(content):
                ht = content[i]
                if ht_name_is(ht, 'br'):
                    ht = ' '
                ht_append(title, ht)
                i += 1

            # Build the new heading.
            e.content = ["Annex " + letter + " ", html.span(status, class_="section-status"), " "] + title
        else:
            # Autogenerate annex subsection number.
            s = letter + "." + ".".join(map(str, annex_counters[1:level + 1])) + '\t'
            if e.content and isinstance(e.content[0], str):
                e.content[0] = s + e.content[0]
            else:
                e.content.insert(0, s)

    tag_names = {
        # ANNEX, a2, a3, a4 are treated specially.
        'Alg2': None,
        'Alg3': None,
        'Alg4': None,
        'Alg40': None,
        'Alg41': None,
        'bibliography': 'li.bibliography-entry',
        'BulletNotlast': 'li',
        'CodeSample3': 'pre',
        'CodeSample4': 'pre',
        'DateTitle': 'h1',
        'ECMAWorkgroup': 'h1',
        'Figuretitle': 'figcaption',
        'Heading1': 'h1',
        'Heading2': 'h1',
        'Heading3': 'h1',
        'Heading4': 'h1',
        'Heading5': 'h1',
        'Introduction': 'h1',
        'M0': None,
        'M4': None,
        'M20': 'div.math-display',
        'MathDefinition4': 'div.display',
        'MathSpecialCase3': 'li',
        'Note': '.Note',
        'RefNorm': 'p.formal-reference',
        'StandardNumber': 'h1',
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

    for e in findall(doc, 'p'):
        num = e.style and '-ooxml-numId' in e.style
        default_tag = 'li' if num else 'p'

        if 'class' in e.attrs:
            cls = e.attrs['class']
            del e.attrs['class']

            if cls not in tag_names:
                unrecognized_styles[cls] += 1
            #e.content.insert(0, html.span('<{0}>'.format(cls), style="color:red"))

            if cls in ('ANNEX', 'a2', 'a3', 'a4'):
                munge_annex_heading(e, cls)
            else:
                tag = tag_names.get(cls)
                if tag is None:
                    tag = default_tag
                elif '.' in tag:
                    tag, _, e.attrs['class'] = tag.partition('.')
                    if tag == '':
                        tag = default_tag
                e.name = tag
        else:
            e.name = default_tag

def fixup_element_spacing(doc):
    """
    Change "A<i> B</i>" to "A <i>B</i>".

    That is, move all start tags to the right of any adjacent whitespace,
    and move all end tags to the left of any adjacent whitespace.
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
            elif k.name == 'pre':
                # Don't mess with spaces in a pre element.
                result.append(k)
            else:
                discard_space = k.is_block()
                if k.content:
                    a = k.content[0]
                    if isinstance(a, str) and a[:1].isspace():
                        k.content[0] = a.lstrip()
                        if not discard_space:
                            addstr(' ')
                result.append(k)
                if k.content:
                    b = k.content[-1]
                    if isinstance(b, str) and b[-1:].isspace():
                        k.content[-1] = b.rstrip()
                        if not discard_space:
                            addstr(' ')

        parent.content[:] = result

    def walk(e):
        rebuild_e = False
        for i, kid in e.kids():
            if (not rebuild_e
                and kid.content
                and ((isinstance(kid.content[0], str) and kid.content[0][:1].isspace())
                     or (isinstance(kid.content[-1], str) and kid.content[-1][-1:].isspace()))):
                # We do not rebuild immediately, but wait until after walking
                # all contents, because if we eject trailing whitespace from
                # the last kid, we want to eject it from the parent too in
                # turn.
                rebuild_e = True
            walk(kid)

        if rebuild_e:
            rebuild(e)

    walk(doc)


def doc_body(doc):
    body = doc.content[1]
    assert body.name == 'body'
    return body

def ht_name_is(ht, name):
    return not isinstance(ht, str) and ht.name == name

def fixup_sec_4_3(doc):
    for parent, i, kid in all_parent_index_child_triples(doc):
        # Hack: Sections 4.3.7 and 4.3.16 are messed up in the document. Wrong style. Fix it.
        if kid.name == "h1" and i > 0 and ht_name_is(parent.content[i - 1], 'h1') and (kid.content == ["built-in object"] or kid.content == ["String value"]):
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

def fixup_sections(doc):
    """ Group h1 elements and subsequent elements of all kinds together into sections. """

    body_elt = doc_body(doc)
    body = body_elt.content

    def starts_with_section_number(s):
        return re.match(r'[1-9]|[A-Z]\.[1-9][0-9]*', s) is not None

    def heading_info(h):
        """ h is an h1 element. Return a pair (sec_num, title).
        sec_num is the section number, as a string, or None.
        title is the title, another string, or None.
        """

        c = h.content
        if len(c) == 0:
            return None, None
        s = c[0]
        if not isinstance(s, str):
            return None, None
        s = s.lstrip()

        # Hack: Section numbers are autogenerated (by Word) for just three of
        # the document's sections. Autogenerate them here.
        if s == 'Scope':
            s = '1\t' + s
        elif s == 'Conformance':
            s = '2\t' + s
        elif s == 'Normative references':
            s = '3\t' + s

        num, tab, title = s.partition('\t')
        if tab == "":
            if len(c) > 1 and ht_name_is(c[1], "span") and c[1].attrs.get("class") == "section-status":
                return s.strip(), ''
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
        if num.startswith('Annex '):
            return num[6:]
        else:
            return num

    def wrap(sec_num, sec_title, start):
        """ Wrap the section starting at body[start] in a section element. """

        sec_id = sec_num_to_id(sec_num) if sec_num else None

        j = start + 1
        while j < len(body):
            kid = body[j]
            if not isinstance(kid, str):
                if kid.name == "div":
                    if (kid.attrs.get("id") == "ecma-disclaimer"
                        or kid.attrs.get("class") == "inner-title"):
                        # Don't let the introduction section eat up these elements.
                        break
                if kid.name == "h1":
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
            c[0:1] = [span, ' ' + sec_title]

        # Actually do the wrapping.
        body[start:stop] = [html.section(*body[start:stop], **attrs)]

    for i, kid in body_elt.kids("h1"):
        num, title = heading_info(kid)
        wrap(num, title, i)

def fixup_hr(doc):
    """ Replace <p><hr></p> with <hr>.

    Word treats an explicit page break as occurring within a paragraph rather
    than between paragraphs, and this leads to goofy markup which has to be
    fixed up.
    """

    for a, i, b in all_parent_index_child_triples(doc):
        if b.name == "p" and len(b.content) == 1 and isinstance(b.content[0], html.Element) and b.content[0].name == "hr":
            a.content[i] = b.content[0]

def fixup_toc(doc):
    """ Generate a table of contents and replace the one in the document. """

    def make_toc_list(e, depth=0):
        sublist = []
        for _, sect in e.kids("section"):
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
        # TODO - make links when there isn't one
        i = 0
        output += h1.content[i:]  # shallow copy, nodes may appear in tree multiple times

        # Find any subsections.
        if depth < 3:
            output += make_toc_list(sect, depth)

        return output

    body = doc_body(doc)
    toc = html.section(html.h1("Contents"), *make_toc_list(body))

    hr_iterator = body.kids("hr")
    i0, first_hr = hr_iterator.__next__()
    i1, next_hr = hr_iterator.__next__()
    body.content[i0: i1 + 1] = [toc]

def fixup_tables(doc):
    """ Turn highlighted td elements into th elements.

    Also, OOXML puts all table cell content in paragraphs; strip out the extra
    <p></p> tags.

    Precedes fixup_code, which converts p elements containing only code into
    pre elements; we don't want code elements in tables to handled that way.
    """
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
            if not p.attrs and not p.style:
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
                if not span.attrs and not span.style:
                    td.content = span.content

def fixup_code(doc):
    """ Merge adjacent code elements. Convert p elements containing only code elements to pre.

    Precedes fixup_notes, which considers pre elements to be part of notes.
    """

    for e in findall(doc, 'p'):
        if len(e.content) == 1 and ht_name_is(e.content[0], 'code'):
            code = e.content[0]
            s = ''
            for k in code.content:
                if isinstance(k, str):
                    s += k
                elif k.name == 'br':
                    s += '\n'
                else:
                    s = None
                    break
            if s is not None:
                e.name = 'pre'
                e.content[:] = [s]

def fixup_notes(doc):
    """ Wrap each NOTE in div.note and wrap the labels "NOTE", "NOTE 2", etc. in span.nh. """

    def find_nh(p, strict=False):
        s = p.content[0]
        if not isinstance(s, str):
            if strict:
                warn("warning in fixup_notes: p.Note paragraph does not start with a string ")
            return None
        else:
            left, tab, right = s.partition('\t')
            if tab is None:
                if strict:
                    warn('warning in fixup_notes: no tab in NOTE: ' + repr(s))
                return None
            elif not left.startswith('NOTE'):
                if strict:
                    warn('warning in fixup_notes: no "NOTE" in p.Note: ' + repr(s))
                return None
            else:
                return left, right

    def can_be_included(next_sibling):
        return (next_sibling.name == 'pre'
                or (next_sibling.name in ('p', 'li')
                    and next_sibling.attrs.get("class") == "Note"
                    and find_nh(next_sibling, strict=False) is None))

    for parent, i, p in all_parent_index_child_triples(doc):
        if p.name == 'p':
            has_note_class = p.attrs.get('class') == 'Note'
            nh_info = find_nh(p, strict=has_note_class)
            if nh_info or has_note_class:
                # This is a note! See if the word "NOTE" or "NOTE 1" can be divided out into
                # a span.nh element. This should ordinarily be the case.
                if nh_info:
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
            if h.content and ht_name_is(h.content[0], 'span') and h.content[0].attrs.get('class') == 'secnum':
                i += 1
            s = ht_text(h.content[i:])
            if s.strip() == title:
                return sect
    raise ValueError("No section has the title " + repr(title))

def fixup_7_9_1(doc, docx):
    """ Fix the ilvl attributes on the bullets in section 7.9.1.

    Precedes fixup_lists which consumes this data.
    """
    bullet_numid = '384' # gross hard-coded hack
    assert docx.get_list_style_at(bullet_numid, '3').numFmt == 'bullet'

    sect = find_section(doc, 'Rules of Automatic Semicolon Insertion')
    lst = sect.content[2:7]
    assert [int(elt.style['-ooxml-ilvl']) for elt in lst] == [2, 0, 0, 2, 2]
    for i in (1, 2):
        lst[i].style['-ooxml-ilvl'] = '3'
        lst[i].style['-ooxml-numId'] = bullet_numid
        assert has_bullet(docx, lst[i])

def fixup_lists(e, docx):
    if e.name in ('ol', 'ul'):
        # This is already a list. I don't think there are any lists we don't
        # identify during transform phase that are nested in lists we do
        # identify. So skip this.
        return

    have_list_items = False
    for _, k in e.kids():
        fixup_lists(k, docx)
        if k.name == 'li':
            have_list_items = True

    if have_list_items:
        # Walk the elements from left to right. If we find any <li> elements,
        # wrap them in <ol> elements to the appropriate depth.
        kids = e.content
        new_content = []
        lists = []
        for k in kids:
            if isinstance(k, str) or k.name != 'li':
                # Not a list item. Close all open lists. Add k to new_content.
                del lists[:]
                new_content.append(k)
            else:
                # Oh no. It is a list item. Well, what is its depth? Does it
                # have a bullet or numbering?
                bullet = False
                if k.style and '-ooxml-ilvl' in k.style:
                    depth = int(k.style['-ooxml-ilvl'])
                    bullet = has_bullet(docx, k)

                    # While we're here, delete the magic style attributes.
                    del k.style['-ooxml-ilvl']
                    del k.style['-ooxml-numId']
                else:
                    depth = 0

                # Close any open lists at greater depth.
                while lists and lists[-1][0] > depth:
                    del lists[-1]

                # If we have a list at equal depth, but it's the wrong kind, close it too.
                if lists and lists[-1][0] == depth and bullet != (lists[-1][1].name == 'ul'):
                    del lists[-1]

                # If we don't already have a list at that depth, open one.
                if not lists or depth > lists[-1][0]:
                    if bullet:
                        new_list = html.ul()
                    else:
                        new_list = html.ol(class_='block' if lists else 'proc')

                    # If there is an enclosing list, add new_list to the last <li>
                    # of the enclosing list, not the enclosing list itself.
                    # If there is no enclosing list, add new_list to new_content.
                    (lists[-1][1].content[-1].content if lists else new_content).append(new_list)
                    lists.append((depth, new_list))

                lists[-1][1].content.append(k)

        kids[:] = new_content

def fixup_list_paragraphs(doc):
    """ Put some more space between list items in certain lists. """

    def is_block(ht):
        return not isinstance(ht, str) and ht.is_block()

    for ul in findall(doc, 'ul'):
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
        if chars / n > 80:
            for _, li in ul.kids('li'):
                i = len(li.content)
                while i > 0 and is_block(li.content[i - 1]):
                    i -= 1
                li.content[:i] = [html.p(*li.content[:i])]

def fixup_picts(doc):
    """ Replace Figure 1 with canned HTML. Remove div.w-pict elements. """
    def walk(e):
        i = 0
        while i < len(e.content):
            child = e.content[i]
            if isinstance(child, str):
                i += 1
            elif child.name == 'div' and child.attrs.get('class') == 'w-pict':
                # Remove the div element, but retain its contents.
                e.content[i:i + 1] = child.content
            elif (child.name == 'p'
                  and len(child.content) == 1
                  and ht_name_is(child.content[0], "div")
                  and child.content[0].attrs.get('class') == 'w-pict'):
                pict = child.content[0]
                is_figure_1 = False
                if i + 1 < len(e.content):
                    caption = e.content[i + 1]
                    if ht_name_is(caption, 'figcaption') and caption.content and caption.content[0].startswith('Figure 1'):
                        is_figure_1 = True

                if is_figure_1:
                    image = html.object(
                        html.img(src="figure-1.png", width="719", height="354", alt="An image of lots of boxes and arrows."),
                        type="image/svg+xml", width="719", height="354", data="figure-1.svg")
                    del e.content[i + 1]
                    e.content[i] = html.figure(image, caption)
                else:
                    # Remove the div element, but retain its contents.
                    e.content[i:i + 1] = pict.content
            else:
                walk(child)
                i += 1

    walk(doc)

def fixup_figures(doc):
    for parent, i, child in all_parent_index_child_triples(doc):
        if child.name == 'figcaption' and i + 1 < len(parent.content) and ht_name_is(parent.content[i + 1], 'figure'):
            # The iterator is actually ok with this mutation, but it's tricky.
            figure = parent.content[i + 1]
            del parent.content[i]
            figure.content.insert(0, child)

def fixup_links(doc):
    sections_by_title = {}
    for sect in findall(doc, 'section'):
        if 'id' in sect.attrs and sect.content and sect.content[0].name == 'h1':
            title = ht_text(sect.content[0].content[1:]).strip()
            sections_by_title[title] = '#' + sect.attrs['id']

    specific_link_source_data = [
        # 5.2
        ("abs(", "Algorithm Conventions"),
        ("sign(", "Algorithm Conventions"),
        ("modulo", "Algorithm Conventions"),
        ("floor(", "Algorithm Conventions"),

        # clause 7
        ("automatic semicolon insertion (7.9)", "Automatic Semicolon Insertion"),
        ("automatic semicolon insertion (see 7.9)", "Automatic Semicolon Insertion"),
        ("semicolon insertion (see 7.9)", "Automatic Semicolon Insertion"),

        # clause 8
        ("Type(", "Types"),
        ("List", "The List and Record Specification Type"),
        ("Completion Record", "The Completion Record Specification Type"),
        ("Completion", "The Completion Record Specification Type"),
        ("NormalValue", "The Completion Record Specification Type"),
        ("NormalCompletion", "The Completion Record Specification Type"),
        ("abrupt completion", "The Completion Record Specification Type"),
        ("ReturnIfAbrupt", "The Completion Record Specification Type"),
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
        ("GetValue", "GetValue (V)"),
        ("PutValue", "PutValue (V, W)"),
        ("Property Descriptor", "The Property Descriptor and Property Identifier Specification Types"),
        ("Property Identifier", "The Property Descriptor and Property Identifier Specification Types"),
        ("IsAccessorDescriptor", "IsAccessorDescriptor ( Desc )"),
        ("IsDataDescriptor", "IsDataDescriptor ( Desc )"),
        ("IsGenericDescriptor", "IsGenericDescriptor ( Desc )"),
        ("FromPropertyDescriptor", "FromPropertyDescriptor ( Desc )"),
        ("ToPropertyDescriptor", "ToPropertyDescriptor ( Obj )"),

        # clause 9
        ("ToPrimitive", "ToPrimitive"),
        ("ToBoolean", "ToBoolean"),
        ("ToNumber", "ToNumber"),
        ("ToInteger", "ToInteger"),
        ("ToInt32", "ToInt32: (Signed 32 Bit Integer)"),
        ("ToUint32", "ToUint32: (Unsigned 32 Bit Integer)"),
        #("ToUint16 (9.7)", "ToUint16: (Unsigned 16 Bit Integer)"),   # flunks the assertion
        ("ToUint16", "ToUint16: (Unsigned 16 Bit Integer)"),
        ("ToString", "ToString"),
        ("ToObject", "ToObject"),
        ("CheckObjectCoercible", "CheckObjectCoercible"),
        ("IsCallable", "IsCallable"),
        ("SameValue (according to 9.12)", "The SameValue Algorithm"),
        ("SameValue", "The SameValue Algorithm"),
        ("the SameValue algorithm (9.12)", "The SameValue Algorithm"),
        ("the SameValue Algorithm (9.12)", "The SameValue Algorithm"),

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
        ("GetIdentifierReference", "GetIdentifierReference (lex, name, strict)"),
        ("NewDeclarativeEnvironment", "NewDeclarativeEnvironment (E)"),
        ("NewObjectEnvironment", "NewObjectEnvironment (O, E)"),
        ("the global environment", "The Global Environment"),
        ("the Global Environment", "The Global Environment"),

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
        ("msPerHour", "Hours, Minutes, Second, and Milliseconds"),
        ("MakeTime", "MakeTime (hour, min, sec, ms)"),
        ("MakeDay", "MakeDay (year, month, date)"),
        ("MakeDate", "MakeDate (day, time)"),
        ("TimeClip", "TimeClip (time)")
    ]

    specific_links = [(text, sections_by_title[title]) for text, title in specific_link_source_data]

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

    SECTION = r'([1-9A-Z][0-9]*(?:\.[1-9][0-9]*)+)'
    def compile(re_source):
        return re.compile(re_source.replace("SECTION", SECTION))

    section_link_regexes = list(map(compile, [
        # Match " (11.1.5)" and " (see 7.9)"
        # The space is to avoid matching "(3.5)" in "Math.round(3.5)".
        r' \(((?:see )?SECTION)\)',

        r'(?:see|See|in|of|to|and) (SECTION)(?:$|\.$|[,:) ]|\.[^0-9])',

        # Match "(Clause 16)", "(see clause 6)".
        r'(?i)(?:)\((?:but )?((?:see\s+(?:also\s+)?)?clause\s+([1-9][0-9]*))\)',

        # Match the first section number in a parenthesized list "(13.3.5, 13.4, 13.6)"
        r'\((SECTION),\ ',

        # Match the first section number in a list at the beginning of a paragraph, "12.14:" or "12.7, 12.7:"
        r'^(SECTION)[,:]',

        # Match the second or subsequent section number in a parenthesized list.
        r', (SECTION)[,):]',

        # Match the penultimate section number in lists that don't use the
        # Oxford comma, like "13.3, 13.4 and 13.5"
        r' (SECTION) and\b',

        # Match "Clause 8" in "as defined in Clause 8 of this specification".
        r'(?i)in (Clause ([1-9][0-9]*))',
    ]))

    # Disallow . ( ) at the end since it's usually not meant as part of the URL.
    url_re = re.compile(r'https?://[0-9A-Za-z;/?:@&=+$,_.!~*()\'-]+[0-9A-Za-z;/?:@&=+$,_!~*\'-]')

    def find_link(s):
        best = None
        for text, target in specific_links:
            i = s.find(text)
            if (i != -1
                and target != current_section  # don't link sections to themselves
                and (i == 0 or not s[i-1].isalnum())  # check for word break before
                and (text.endswith('(')
                     or i + len(text) == len(s)
                     or not s[i + len(text)].isalnum())  # and after
                and (best is None or i < best[0])):
                # New best hit.
                n = len(text)
                if text.endswith('('):
                    n -= 1
                best = i, i + n, target

        for link_re in section_link_regexes:
            m = link_re.search(s)
            while m is not None:
                id = "sec-" + m.group(2)
                if id not in all_ids:
                    warn("no such section: " + m.group(2))
                    m = link_re.search(s, m.end(1))
                else:
                    hit = m.start(1), m.end(1), "#sec-" + m.group(2)
                    if best is None or hit < best:
                        best = hit
                    break

        m = url_re.search(s)
        if m is not None:
            hit = m.start(), m.end(), m.group(0)
            if best is None or hit < best:
                best = hit

        return best

    def linkify(parent, i, s):
        while True:
            m = find_link(s)
            if m is None:
                return
            start, stop, href = m
            if start > 0:
                parent.content.insert(i, s[:start])
                i += 1
            assert not href.startswith('#') or href[1:] in all_ids
            parent.content[i] = html.a(href=href, *s[start:stop])
            i += 1
            if stop < len(s):
                parent.content.insert(i, s[stop:])
            else:
                break
            s = s[stop:]

    current_section = None
    def visit(e):
        nonlocal current_section

        id = e.attrs.get('id')
        if id is not None:
            current_section = '#' + id

        for i, kid in enumerate(e.content):
            if isinstance(kid, str):
                linkify(e, i, kid)
            elif kid.name == 'a' and 'href' in kid.attrs:
                # Yo dawg. No links in links.
                pass
            elif kid.name == 'h1' or (kid.name == 'ol' and kid.attrs.get('class') == 'toc'):
                # Don't linkify headings or the table of contents.
                pass
            else:
                visit(kid)

        if id is not None:
            current_section = None

    visit(doc_body(doc))

def fixup_remove_hr(doc):
    """ Remove all remaining hr elements. """
    for parent, i, child in all_parent_index_child_triples(doc):
        if child.name == 'hr':
            del parent.content[i]

def fixup_title_page(doc):
    """ Apply a handful of fixups to the junk gleaned from the title page. """
    for parent, i, child in all_parent_index_child_triples(doc):
        if parent.name == 'p' and child.name == 'h1':
            # A p element shouldn't contain an h1, so make this an hgroup.
            parent.name = 'hgroup'
            if len(parent.content) != 6:
                continue
            h = parent.content[1]

            # One of the lines has an ugly typo that I don't want right up
            # front in large type.
            s = h.content[-1]
            assert s.endswith(' , 2012')
            h.content[-1] = s.replace(' , 2012', ', 2012')

            # A few of the lines here are redundant.
            del parent.content[3:]

def fixup_overview_biblio(doc):
    sect = find_section(doc, "Overview")
    for i, p in enumerate(sect.content):
        if p.name == 'p':
            if p.content and p.content[0].startswith('Gosling'):
                break

    # First, strip the <sup> element around the &trade; symbol.
    assert p.content[1].name == 'sup'
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

def fixup_grammar_pre(doc):
    """ Convert runs of div.lhs and div.rhs elements in doc to pre elements.

    Keep the text; throw everything else away.
    """

    def is_grammar(div):
        return ht_name_is(div, 'div') and div.attrs.get('class') in ('lhs', 'rhs')

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

    for parent, i, div in all_parent_index_child_triples(doc):
        if is_grammar(div):
            j = i + 1
            while j < len(parent.content) and is_grammar(parent.content[j]):
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

            # Hack - not all the paragraphs marked as syntax are actually
            # things we want to replace. So as a heuristic, only make the
            # change if the first line satisfied is_lhs.
            if syntax.startswith('    '):
                continue

            parent.content[i:j] = [html.pre(syntax, class_="syntax")]

def fixup_grammar_post(doc):
    """ Generate nice markup from the stripped-down pre.syntax elements
    created by fixup_grammar_pre. """

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
        | \[lookahead \  . [^]]* \]     # the . stands for &notin;
        | <[A-Z]+>                      # special character
        | \(                            # unstick a parenthesis from the following token
        | [^ ]*                         # any other token
        )\s*
        ''')

    def markup_div(text, cls, xrefs=None):
        xref = None
        markup = []

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
            elif token and token.rstrip(':') == '':
                markup.append(html.span(token, class_='geq'))
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

    for parent, i, pre in all_parent_index_child_triples(doc):
        if pre.name == 'pre' and pre.attrs.get('class') == 'syntax':
            divs = []
            [syntax] = pre.content
            syntax = syntax.lstrip('\n')
            for production in syntax.split('\n\n'):
                lines = production.splitlines()

                assert not lines[0][:1].isspace()
                lines_out = markup_div(lines[0], 'lhs')
                for line in lines[1:]:
                    assert line.startswith('    ')
                    lines_out += markup_div(line.strip(), 'rhs')
                divs.append(html.div(*lines_out, class_='gp'))
            parent.content[i:i+1] = divs

def fixup_add_disclaimer(doc):
    div = html.div
    p = html.p
    strong = html.strong
    em = html.em
    a = html.a

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
          "It has doubtlessly stripped out or garbled some of the formatting that makes "
          "the specification comprehensible. You can help improve the program ",
          a("here", href="https://github.com/jorendorff/es-spec-html"),
          "."),
        # (U+2019 is RIGHT SINGLE QUOTATION MARK, the character you're supposed to use for an apostrophe.)
        p("For copyright information, see ECMA\u2019s legal disclaimer in the document itself."),
        id="unofficial")
    doc_body(doc).content.insert(0, disclaimer)

def fixup(docx, doc):
    styles = docx.styles
    numbering = docx.numbering

    fixup_list_styles(doc, docx)
    fixup_formatting(doc, styles)
    fixup_paragraph_classes(doc)
    fixup_element_spacing(doc)
    fixup_sec_4_3(doc)
    fixup_sections(doc)
    fixup_hr(doc)
    fixup_toc(doc)
    fixup_tables(doc)
    fixup_code(doc)
    fixup_notes(doc)
    fixup_7_9_1(doc, docx)
    fixup_lists(doc, docx)
    fixup_list_paragraphs(doc)
    fixup_picts(doc)
    fixup_figures(doc)
    fixup_links(doc)
    fixup_remove_hr(doc)
    fixup_title_page(doc)
    fixup_overview_biblio(doc)
    fixup_grammar_pre(doc)
    fixup_grammar_post(doc)
    fixup_add_disclaimer(doc)
    return doc
