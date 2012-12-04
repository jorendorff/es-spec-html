import re
import html as _html
from html.entities import codepoint2name as _entities
from io import StringIO
import textwrap

class Element:
    __slots__ = ['name', 'attrs', 'style', 'content']
    def __init__(self, name, attrs, style, content):
        assert isinstance(name, str)
        self.name = name
        assert attrs is None or isinstance(attrs, dict)
        self.attrs = attrs or {}
        assert style is None or isinstance(style, dict)
        self.style = style
        assert not isinstance(content, str)
        self.content = list(content)
        assert all(isinstance(item, (str, Element)) for item in self.content)

    def is_block(self):
        """ True if this is a block element.

        Block elements have the property that extra whitespace before or after
        either the start tag or the end tag is ignored. (Some elements, like
        <section>, are expected to contain only block elements and no text; but
        of course there are also block elements, like <p>, that contain inline
        elements and text.)"""

        return _tag_data[self.name][0] == 'B'

    def to_html(self):
        f = StringIO()
        write_html(f, self, strict=False)
        return f.getvalue()

    __repr__ = to_html

    def kids(self, name=None):
        """ Return an iterator over (index, child) pairs where each child is a
        child element of self. The iterator is robust against mutating
        self.content to the right of i. """
        for i, kid in enumerate(self.content):
            if isinstance(kid, Element) and (name is None or kid.name == name):
                yield i, kid

    def with_content(self, replaced_content):
        """ Return a copy of self with different content. """
        return Element(self.name, self.attrs, self.style, replaced_content)

    def with_content_slice(self, start, stop, replaced_content):
        copy = self.content[:]
        copy[start:stop] = replaced_content
        return self.with_content(copy)

    def with_(self, name=None, attrs=None, style=None, content=None):
        if name is None: name = self.name
        if attrs is None: attrs = self.attrs
        if style is None: style = self.style
        if content is None: content = self.content
        return Element(name, attrs, style, content)

    def find_replace(self, matcher, replacement):
        """ A sort of map() on htmodel content.

        self - An Element to transform.

        matcher - A function mapping an Element to True, False, or None.

        replacement - A function mapping an Element to a content list.

        Walk the entire tree under the Element self; for each element e such
        that matcher(e) is True, call replacement(e); return a tree with each
        such Element replaced by the content in replacement(e).

        If matcher(e) is False or None, replacement(e) is not called, and the
        element is left in the result tree; the difference is that if
        matcher(e) is False, e's descendants are visited; if it is None, the
        descendants are skipped entirely and left unchanged in the result tree.

        self is left unmodified, but the result is not a deep copy: it may be
        self or an Element whose tree shares some parts of self.
        """
        def map_element(e):
            match_result = matcher(e)
            assert match_result is True or match_result is False or match_result is None
            if match_result is None:
                return [e]
            replaced_content = map_content(e.content)
            if replaced_content is e.content:
                e2 = e
            else:
                e2 = e.with_content(replaced_content)

            if match_result:
                return list(replacement(e2))
            else:
                return [e2]

        def map_content(source):
            changed = False
            result = []
            for child in source:
                if isinstance(child, str):
                    result.append(child)
                else:
                    seq = map_element(child)
                    changed = changed or seq != [child]
                    result += seq
            if changed:
                return result
            else:
                return source

        result_content = map_element(self)
        if len(result_content) != 1:
            raise ValueError("replaced root element with {} pieces of content".format(len(result_content)))
        result_elt = result_content[0]
        if not isinstance(result_elt, Element):
            raise ValueError("replaced root element with non-element content")
        return result_elt


    def replace(self, name, replacement):
        """ A sort of map() on htmodel content.

        self - An Element to transform.

        name - A string.

        replacement - A function taking a single Element and returning a content list
            (that is, a list of Elements and/or strings).

        Walk the entire tree under the Element self; for each Element e with the given
        name, call replacement(e); return a tree with each such Element replaced by
        the content in replacement(e).

        If self.name == name and list(replacement(self)) is not a list consisting of
        exactly one Element, raise a ValueError.

        self is left unmodified, but the result is not a deep copy: it may be
        self or an Element whose tree shares some parts of self.
        """

        return self.find_replace(lambda e: e.name == name, replacement)

def escape(s, quote=False):
    def replace(m):
        c = ord(m.group(0))
        if c in _entities:
            return '&' + _entities[c] + ';'
        return '&#x{:x};'.format(c)

    # The stdlib takes care of & > < ' " for us.
    s = _html.escape(s, quote)

    # Now we only need to worry about non-ascii characters.
    return re.sub('[^\n -~]', replace, s)

non_indenting_tags = {'html', 'body'}

def save_html(filename, ht):
    assert ht.name == 'html'

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("<!doctype html>\n")
        write_html(f, ht)

def write_html(f, ht, indent='', strict=True):
    WIDTH = 130

    def htmlify(name):
        """ Convert a pythonified tag name or attribute name back to HTML. """
        if name.endswith('_'):
            name = name[:-1]
        name = name.replace('_', '-')
        return name

    def start_tag(ht):
        attrs = ''.join(' {0}="{1}"'.format(htmlify(k), escape(v, True))
                        for k, v in ht.attrs.items())
        if ht.style:
            assert 'style' not in ht.attrs
            style = '; '.join(name + ": " + value for name, value in sorted(ht.style.items()))
            attrs += ' style="{0}"'.format(style)
        return '<{0}{1}>'.format(ht.name, attrs)

    def is_ht_inline(ht):
        return isinstance(ht, str) or _tag_data[ht.name][0] == 'I'

    def wrap_text(text, indent, subsequent_indent):
        """ Add indentation and newlines to text.

        The result is either empty or ends with a newline.
        """
        result = textwrap.fill(text, WIDTH, expand_tabs=False,
                               replace_whitespace=False, initial_indent=indent,
                               subsequent_indent=subsequent_indent, break_long_words=False,
                               break_on_hyphens=False)
        if result:
            result += '\n'
        return result

    def content_to_str(content, indent):
        """ Convert the given list of content to HTML.
        
        If all the content is inline, return a string containing no newlines
        and no leading indentation (to be indented word-wrapped by the
        caller). Otherwise return a string ending with a newline, with every
        line indented.
        """
        kids = [ht_to_str(ht, indent) for ht in content]
        if any(s.endswith('\n') for s in kids):
            # Block content. This is nontrivial because we must convert
            # any inline children to word-wrapped paragraphs.
            result = ''
            acc = ''
            prev_needs_space = False
            for kid, s in zip(content, kids):
                needs_space = False
                if s.endswith('\n'):
                    # Ridiculous heuristic to determine if there should be a blank line
                    # between this element and the previous one.
                    needs_space = (isinstance(kid, Element)
                                   and (kid.name in {'p', 'figure', 'section'}
                                        or (kid.name in {'div', 'li', 'td'}
                                            and kid.content and not is_ht_inline(kid.content[0]))))

                    if acc:
                        result += wrap_text(acc, indent, indent)
                        acc = ''
                    elif result and (needs_space or prev_needs_space):
                            result += '\n'  # blank line between blocks
                    result += s
                else:
                    acc += s
                prev_needs_space = needs_space
            if acc:
                result += wrap_text(acc, indent, indent)
            return result
        else:
            # All inline content.
            return ''.join(kids)

    def element_type_is_empty(tag_name):
        return _tag_data[tag_name][1] == '0'

    def element_renders_as_block(tag_name):
        return _tag_data[tag_name][0] == 'B'

    def ht_to_str(ht, indent):
        if isinstance(ht, str):
            result = escape(ht)
            if '\n' in ht:
                # This is unexpected; don't word-wrap. Write it nearly-verbatim, with newlines.
                result = escape(result)
                if not result.endswith('\n'):
                    result += '\n'
            return result
        else:
            assert isinstance(ht, Element)
            empty_type = element_type_is_empty(ht.name)
            assert len(ht.content) == 0 or not empty_type  # empty element is empty

            start = start_tag(ht)
            
            if empty_type and len(ht.content) == 0:
                end = ''
            else:
                end = '</{}>'.format(ht.name)

            inner_indent = indent
            if ht.name not in non_indenting_tags:
                inner_indent += '  '
            inner = content_to_str(ht.content, inner_indent)

            if inner.endswith('\n'):
                return (indent + start + '\n' + inner + indent + end + '\n')
            elif element_renders_as_block(ht.name):
                # All inline content. Add start and end tags, then word-wrap it.
                subsequent_indent = indent
                if ht.name != 'p':
                    subsequent_indent += '    '
                return wrap_text(start + inner + end, indent, subsequent_indent)
            else:
                return start + inner + end

    f.write(ht_to_str(ht, indent))

_tag_data = {}

def _init(v):
    # Every tag is one of:
    #
    # - like section, table, tr, ol, ul: block containing only blocks
    #
    # - like p, h1: block containing only inline content
    #
    # - like span, em, strong, i, b: inline containing only inline content
    #
    # - like li or td: block containing either block or inline content (or in the
    #   unique case of li, inline followed by a list)
    #
    # - like hr, img: block containing nothing
    #
    # - like br, wbr: inline containing nothing
    #
    # - like style, script: block with unparsed content
    #
    # - like svg, foreignObject: don't even try to figure out where these should appear
    #
    # Of course 
    _tag_raw_data = '''\
    html head body section hgroup table tbody thead tfoot tr ol\
        ul blockquote ol ul dl figure: Bb
    title p h1 h2 h3 h4 h5 h6 address figcaption pre: Bi
    a em strong small s cite q dfn abbr data time code var samp\
        kbd sub sup i b u mark bdi bdo span: Ii
    div li td th noscript object dt dd: B?
    meta link hr img: B0
    br wbr: I0
    style script: Bs
    svg foreignObject: Ib'''

    def element_constructor(name):
        def construct(*content, **attrs):
            if 'class_' in attrs:
                attrs['class'] = attrs['class_']
                del attrs['class_']
            return Element(name, attrs, None, list(content))
        construct.__name__ = name
        return construct

    for line in _tag_raw_data.splitlines():
        names, colon, info = line.partition(':')
        assert colon
        info = info.strip()
        for name in names.split():
            v[name] = element_constructor(name)
            _tag_data[name] = info

_init(vars())

__all__ = list(_tag_data.keys())
