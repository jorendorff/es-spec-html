import htmodel as html

## def all_elements(e):
##     yield e
##     for k in e.content:
##         if not isinstance(k, str):
##             for d in all_elements(k):
##                 yield d

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
                        or kid.attrs.get("class_") == "inner-title"):
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

    i = 0
    while i < len(body):
        kid = body[i]
        if not isinstance(kid, str) and kid.name == "h1":
            num, title = heading_info(kid)
            wrap(num, title, i)
        i += 1

def fixup_code(e):
    """ Merge adjacent code elements. Convert p elements containing only code elements to pre. """

    i = 0
    while i < len(e.content):
        k = e.content[i]
        if isinstance(k, str):
            pass
        elif k.name == 'code':
            while i + 1 < len(e.content) and not isinstance(e.content[i + 1], str) and e.content[i + 1].name == 'code':
                # merge two adjacent code nodes (should merge adjacent text nodes after this :-P)
                k.content += e.content[i + 1].content
                del e.content[i + 1]
        else:
            fixup_code(k)
        i += 1

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
                and next_sibling.attrs.get("class_") == "note"
                and find_nh(next_sibling) is None)

    def fixup_notes_in(e):
        i = 0
        while i < len(e.content):
            k = e.content[i]
            if isinstance(k, str):
                pass
            elif k.name == 'div' and k.attrs.get('class_') == 'note':
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

            i += 1

    fixup_notes_in(e)

def fixup_lists(e):
    if e.name in ('ol', 'ul'):
        # This is already a list. I don't think there are any lists we don't
        # identify during transform phase that are nested in lists we do
        # identify. So skip this.
        return

    kids = e.content

    have_list_items = False
    for k in kids:
        if not isinstance(k, str):
            fixup_lists(k)
            if k.name == 'li':
                have_list_items = True

    # Walk the elements from left to right. If we find any <li> elements,
    # wrap them in <ol> elements to the appropriate depth.
    new_content = []
    lists = []
    for k in kids:
        if isinstance(k, str) or k.name != 'li':
            # Not a list item. Close all open lists. Add k to new_content.
            del lists[:]
            new_content.append(k)
        else:
            # Oh no. It is a list item. Well, what is its depth?
            if k.style and '@num' in k.style:
                depth = int(k.style['@num'].partition('/')[0])

                # While we're here, delete the @num magic style attribute.
                del k.style['@num']
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
            if isinstance(sib, str) or sib.name != "div" or sib.attrs.get("class_") != "rhs":
                break
            j += 1

        if j > start + 1:
            stop = j
            parent_list[start].attrs["class_"] = "lhs"
            parent_list[start:stop] = [html.div(*parent_list[start:stop], class_="gp")]

    i = 0
    while i < len(e.content):
        kid = e.content[i]
        if not isinstance(kid, str):
            if kid.name == "div" and kid.attrs.get("class_") == "gp":
                wrap(e.content, i)
            elif kid.name in ('body', 'section', 'div'):
                fixup_grammar(kid)
        i += 1


def fixup(doc):
    fixup_sections(doc)
    fixup_code(doc)
    fixup_notes(doc)
    fixup_lists(doc)
    fixup_grammar(doc)
    return doc
