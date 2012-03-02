import htmodel as html
import collections

def all_parent_index_child_triples(e):
    for i, k in enumerate(e.content):
        if not isinstance(k, str):
            yield e, i, k
            for t in all_parent_index_child_triples(k):
                yield t

def fixup_formatting(doc):
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

    def rewrite_adjacent_spans(parent, i, j):
        spans = parent.content[i:j]  # copies a slice of the array

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
            set_current_style_to(kid.style)
            all_content += kid.content
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

        parent.content[i:j] = build_result(ranges, 0, len(all_content))

    for parent, i, kid in all_parent_index_child_triples(doc):
        if kid.name == 'span':
            # We assert the spans don't have attrs because we are going to
            # rewrite these guys retaining only the style. This fixup needs to
            # happen early enough in rewriting that this isn't a problem; it also
            # has to be early so that other markup doesn't get in the way.
            assert not kid.attrs
            j = i + 1
            while j < len(parent.content):
                sibling = parent.content[j]
                if isinstance(sibling, str) or sibling.name != 'span':
                    break
                assert not sibling.attrs
                j += 1

            # Call rewrite even if j == i + 1, to make sure new_span gets called,
            # converting, for example
            #   <span style="font-style: italic">foo</span>
            # to
            #   <i>foo</i>
            rewrite_adjacent_spans(parent, i, j)

def findall(e, name):
    if e.name == name:
        yield e
    for k in e.content:
        if not isinstance(k, str):
            for d in findall(k, name):
                yield d

unrecognized_styles = collections.defaultdict(int)

def fixup_paragraph_classes(doc):
    tag_names = {
        'ANNEX': 'h1',
        'Alg2': None,
        'Alg3': None,
        'Alg4': None,
        'Alg40': None,
        'Alg41': None,
        'bibliography': 'li.bibliography-entry',
        'BulletNotlast': 'li',
        'CodeSample3': 'pre',
        'CodeSample4': 'pre',
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
                #e.content.insert(0, span('<{0}>'.format(cls), style="color:red"))

            if cls == 'Note':
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

    body = doc.content[1]
    assert body.name == 'body'
    everything = html.section(disclaimer, id="everything")
    everything.content += body.content
    body.content = [everything]

def fixup_sec_4_3(doc):
    for parent, i, kid in all_parent_index_child_triples(doc):
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

        num, tab, title = s.partition('\t')
        if tab is None:
            parts = s.split(None, 1)
            if len(parts) == 2:
                num, title = parts
            else:
                return None, s

        return num, title

    def contains(a, b):
        """ True if section `a` contains section `b` as a subsection.
        `a` and `b` are section numbers, which are strings; but some sections
        do not have numbers, so either or both may be None.
        """
        return a is not None and (b is None or b.startswith(a + "."))

    def wrap(sec_num, sec_title, start):
        """ Wrap the section starting at body[start] in a section element. """
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
                    if contains(sec_num, kid_num):
                        # kid starts a subsection. Wrap it!
                        wrap(kid_num, kid_title, j)
                    else:
                        # kid starts the next section. Done!
                        break
            j += 1
        stop = j

        attrs = {}
        if sec_num is not None and sec_num[:1].isdigit():
            attrs['id'] = "sec-" + sec_num
            span = html.span(
                html.a(sec_num, href="#sec-" + num, title="link to this section"),
                class_="secnum")
            c = body[start].content
            c[0] = ' ' + sec_title
            c.insert(0, span)

        # Actually do the wrapping.
        body[start:stop] = [html.section(*body[start:stop], **attrs)]

    assert len(doc.content) == 2
    body_elt = doc.content[1]
    assert body_elt.name == "body"
    everything_elt = body_elt.content[0]
    assert everything_elt.attrs['id'] == "everything"

    body = everything_elt.content

    for i, kid in everything_elt.kids("h1"):
        num, title = heading_info(kid)
        wrap(num, title, i)

def fixup_code(e):
    """ Merge adjacent code elements. Convert p elements containing only code elements to pre. """

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

def fixup_lists(e):
    if e.name in ('ol', 'ul'):
        # This is already a list. I don't think there are any lists we don't
        # identify during transform phase that are nested in lists we do
        # identify. So skip this.
        return

    have_list_items = False
    for _, k in e.kids():
        fixup_lists(k)
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
                # Oh no. It is a list item. Well, what is its depth?
                if k.style and '-ooxml-ilvl' in k.style:
                    depth = int(k.style['-ooxml-ilvl'])

                    # While we're here, delete the magic style attributes.
                    del k.style['-ooxml-ilvl']
                    del k.style['-ooxml-numId']
                else:
                    depth = 0

                # Close any open lists at greater depth.
                while lists and lists[-1][0] > depth:
                    del lists[-1]

                # If we don't already have a list at that depth, open one.
                if not lists or depth > lists[-1][0]:
                    new_list = html.ol(class_='block' if lists else 'proc')

                    # If there is an enclosing list, add new_list to the last <li>
                    # of the enclosing list, not the enclosing list itself.
                    # If there is no enclosing list, add new_list to new_content.
                    (lists[-1][1][-1].content if lists else new_content).append(new_list)
                    lists.append((depth, new_list.content))

                lists[-1][1].append(k)

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

def fixup_hr(doc):
    """ Replace <p><hr></p> with <hr>.

    Word treats an explicit page break as occurring within a paragraph rather
    than between paragraphs, and this leads to goofy markup which has to be
    fixed up.
    """
    for a, i, b in all_parent_index_child_triples(doc):
        if b.name == "p" and len(b.content) == 1 and isinstance(b.content[0], html.Element) and b.content[0].name == "hr":
            a.content[i] = b.content[0]

def fixup(doc):
    fixup_formatting(doc)
    fixup_paragraph_classes(doc)
    fixup_element_spacing(doc)
    insert_disclaimer(doc)
    fixup_sec_4_3(doc)
    fixup_sections(doc)
    fixup_code(doc)
    fixup_notes(doc)
    fixup_lists(doc)
    fixup_grammar(doc)
    fixup_hr(doc)
    return doc
