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
        'Note': 'div.note',  # TODO needs default_tag *also*
        'RefNorm': 'p.formal-reference',
        'StandardNumber': 'h1',
        'StandardTitle': 'h1',
        'Syntax': 'h2',
        'SyntaxDefinition': 'div.rhs',
        'SyntaxDefinition2': 'div.rhs',
        'SyntaxRule': 'div.gp',
        'SyntaxRule2': 'div.gp',
        'Tabletitle': 'figcaption',
        'TermNum': 'h1',
        'Terms': 'p.Terms',  # TODO needs later fixup
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
            elif cls == 'Note':
                # Total special case. Wrap in div.note.
                e.name = 'div'
                e.content = [html.Element(default_tag, e.attrs, e.style, e.content)]
                e.style = None
                e.attrs = {'class': 'note'}
            else:
                tag = tag_names.get(cls)
                if tag is None:
                    tag = default_tag
                elif '.' in tag:
                    tag, _, e.attrs['class'] = tag.partition('.')
                e.name = tag
        else:
            e.name = default_tag

def fixup_element_spacing(doc):
    """
    Change "A<i> B</i>" to "A <i>B</i>".

    That is, move all start tags to the right of any adjacent whitespace,
    and move all end tags to the left of any adjacent whitespace.
    """

    block_elements = {
        'html', 'body', 'p', 'div', 'section',
        'table', 'tbody', 'tr', 'th', 'td', 'li', 'ol', 'ul'
    }

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
                discard_space = k.name in block_elements
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

def insert_disclaimer(doc):
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
    doc_body(doc).content.insert(0, disclaimer)

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
        if len(td.content) == 1 and ht_name_is(td.content[0], 'p'):
            p = td.content[0]
            if p.style and p.style.get('background-color') == '#C0C0C0':
                td.name = 'th'
                del p.style['background-color']
            if len(p.content) == 1 and ht_name_is(p.content[0], 'span'):
                span = p.content[0]
                if span.style and span.style.get('background-color') == '#C0C0C0':
                    td.name = 'th'
                    del span.style['background-color']

            # If the p is vacuous, kill it.
            if not p.attrs and not p.style:
                td.content = p.content

            # Ditto if it happens to contain an empty span.
            if len(td.content) == 1 and ht_name_is(td.content[0], 'span'):
                span = td.content[0]
                if not span.attrs and not span.style:
                    td.content = span.content

def fixup_code(e):
    """ Merge adjacent code elements. Convert p elements containing only code elements to pre.

    Precedes fixup_notes, which considers pre elements to be part of notes.
    """

    for i, k in e.kids():
        if k.name == 'code':
            while i + 1 < len(e.content) and not isinstance(e.content[i + 1], str) and e.content[i + 1].name == 'code':
                # merge two adjacent code nodes (should merge adjacent text nodes after this :-P)
                k.content += e.content[i + 1].content
                del e.content[i + 1]
        else:
            fixup_code(k)

    if e.name == 'p' and len(e.content) == 1 and not isinstance(e.content[0], str) and e.content[0].name == 'code':
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

def fixup_notes(e):
    """ Apply several fixes to div.note elements throughout the document. """

    def find_nh(div, verbose=False):
        first_para = div.content[0]
        s = first_para.content[0]
        if not isinstance(s, str):
            if verbose: print("warning in fixup_notes: div.note paragraph does not start with a string")
            return None
        else:
            left, tab, right = s.partition('\t')
            if tab is None:
                if verbose: print('warning in fixup_notes: no tab in div.note:', repr(s))
                return None
            elif not left.startswith('NOTE'):
                if verbose: print('warning in fixup_notes: no "NOTE" in div.note:', repr(s))
                return None
            else:
                return first_para, left, right

    def can_be_eaten(next_sibling):
        if isinstance(next_sibling, str):
            return False
        return next_sibling.name == 'pre'

    def can_be_combined(next_sibling):
        return (not isinstance(next_sibling, str)
                and next_sibling.name == 'div'
                and next_sibling.attrs.get("class") == "note"
                and find_nh(next_sibling) is None)

    def fixup_notes_in(e):
        for i, k in e.kids():
            if k.name == 'div' and k.attrs.get('class') == 'note':
                # This is a note! See if the word "NOTE" or "NOTE 1" can be divided out into
                # a span.nh element. This should ordinarily be the case.
                nh_info = find_nh(k, verbose=True)
                if nh_info is not None:
                    container, nh, rest = nh_info
                    assert container.content[0] == nh + '\t' + rest
                    container.content[0] = ' ' + rest
                    container.content.insert(0, html.span(nh, class_="nh"))

                # Now look for neighboring div.note elements. These normally are
                # other paragraphs belonging to the same note--in which case, merge
                # them.
                while i + 1 < len(e.content):
                    next_sibling = e.content[i + 1]
                    if can_be_eaten(next_sibling):
                        k.content.append(next_sibling)
                        del e.content[i + 1]
                    elif can_be_combined(next_sibling):
                        k.content += next_sibling.content
                        del e.content[i + 1]
                    else:
                        break
            else:
                # Recurse.
                fixup_notes(k)

    fixup_notes_in(e)

def find_section(doc, sec_id):
    # super slow algorithm
    want_id = "sec-" + sec_id
    for sect in findall(doc, 'section'):
        id = sect.attrs.get('id')
        if id == want_id:
            return sect
    return None

def fixup_7_9_1(doc, docx):
    """ Fix the ilvl attributes on the bullets in section 7.9.1.

    Precedes fixup_lists which consumes this data.
    """
    bullet_numid = '384' # gross hard-coded hack
    assert docx.get_list_style_at(bullet_numid, '3').numFmt == 'bullet'

    sect = find_section(doc, '7.9.1')
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

def fixup_grammar(e):
    """ Fix div.gp and div.rhs elements in `e` by changing the div.gp to a div.lhs
    and wrapping them both in a new div.gp. """

    def wrap(parent_list, start):
        j = start + 1
        while j < len(parent_list):
            sib = parent_list[j]
            if isinstance(sib, str) or sib.name != "div" or sib.attrs.get("class") != "rhs":
                break
            j += 1

        if j > start + 1:
            stop = j
            parent_list[start].attrs["class"] = "lhs"
            parent_list[start:stop] = [html.div(*parent_list[start:stop], class_="gp")]

    for i, kid in e.kids():
        if kid.name == "div" and kid.attrs.get("class") == "gp":
            wrap(e.content, i)
        elif kid.name in ('body', 'section', 'div'):
            fixup_grammar(kid)

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
    all_ids = set([kid.attrs['id'] for _, _, kid in all_parent_index_child_triples(doc) if 'id' in kid.attrs])

    SECTION = r'([1-9A-Z][0-9]*(?:\.[1-9][0-9]*)+)'
    def compile(re_source):
        return re.compile(re_source.replace("SECTION", SECTION))

    section_link_regexes = list(map(compile, [
        # Match " (11.1.5)" and " (see 7.9)"
        # The space is to avoid matching "(3.5)" in "Math.round(3.5)".
        r' \(((?:see )?SECTION)\)',

        # Match "(Clause 16)", "(see clause 6)".
        r'(?i)(?:)\((?:but )?((?:see\s+(?:also\s+)?)?clause\s+([1-9][0-9]*))\)',

        # Match the first section number in a parenthesized list "(13.3.5, 13.4, 13.6)"
        r'\((SECTION),\ ',

        # Match the first section number in a list at the beginning of a paragraph, "12.14:" or "12.7, 12.7:"
        r'^(SECTION)[,:]',

        # Match the second or subsequent section number in a parenthesized list.
        r', (SECTION)[,):]',

        # Match "Clause 8" in "as defined in Clause 8 of this specification".
        r'(?i)in (Clause ([1-9][0-9]*))',

        # Match 
        r'in ((\b[1-9A-Z][0-9]*(?:\.[1-9][0-9]*)+))'
    ]))

    def find_link(s):
        best = None
        for link_re in section_link_regexes:
            m = link_re.search(s)
            while m is not None:
                id = "sec-" + m.group(2)
                if id not in all_ids:
                    warn("no such section: " + m.group(2))
                    m = link_re.search(s, m.end(1))
                    continue
                hit = m.start(1), m.end(1), "#sec-" + m.group(2)
                if best is None or hit < best:
                    best = hit
                break
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

    def visit(e):
        for i, kid in enumerate(e.content):
            if isinstance(kid, str):
                linkify(e, i, kid)
            else:
                visit(kid)

    visit(doc_body(doc))

def fixup(docx, doc):
    styles = docx.styles
    numbering = docx.numbering

    fixup_list_styles(doc, docx)
    fixup_formatting(doc, styles)
    fixup_paragraph_classes(doc)
    fixup_element_spacing(doc)
    insert_disclaimer(doc)
    fixup_sec_4_3(doc)
    fixup_sections(doc)
    fixup_hr(doc)
    fixup_toc(doc)
    fixup_tables(doc)
    fixup_code(doc)
    fixup_notes(doc)
    fixup_7_9_1(doc, docx)
    fixup_lists(doc, docx)
    fixup_grammar(doc)
    fixup_picts(doc)
    fixup_figures(doc)
    fixup_links(doc)
    return doc
