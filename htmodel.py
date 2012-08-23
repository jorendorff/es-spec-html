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
        self.content = list(content)

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

    def replace(self, name, replacement):
        """ A sort of map() on htmodel content.

        self - An Element to transform.

        name - A string.

        replacement - A function taking a single Element and returning a content list
            (that is, a list of Elements and/or strings).

        Walk the entire tree under the Element self; for each Element e with the given
        name, call replacement(e); return a tree with each such Element replaced by
        the elements in replacement(e).

        If self.name == name and list(replacement(self)) is not a list consisting of
        exactly one Element, raise a ValueError.

        self is left unmodified, but the result is not a deep copy: it may be
        self or an Element whose tree shares some parts of self.
        """

        def map_element(e):
            replaced_content = map_content(e.content)
            if replaced_content is not e.content:
                e = e.with_content(replaced_content)
            if e.name == name:
                return list(replacement(e))
            return [e]

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

empty_tags = {'meta', 'br', 'hr', 'link'}
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

    def write_inline_content(f, content, indent, allow_last_block, strict):
        if allow_last_block and isinstance(content[-1], Element) and content[-1].name in ('ol', 'ul', 'table'):
            last = content[-1]
            content = content[:-1]
        else:
            last = None

        for kid in content:
            if isinstance(kid, str):
                f.write(escape(kid))
            else:
                if strict and not is_ht_inline(kid):
                    raise ValueError("block element <{}> can't appear in inline content".format(kid.name))
                write_html(f, kid, indent, strict)

        if last is not None:
            f.write('\n')
            write_html(f, last, indent, strict)
            f.write(indent[:-2])

    if isinstance(ht, str):
        assert not strict
        f.write(escape(ht))
        return

    info = _tag_data[ht.name]
    content = ht.content
    assert info[1] != '0' or len(content) == 0  # empty tag is empty

    if (ht.name in ('p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6')
          or (ht.name == 'li' and ht.content and is_ht_inline(ht.content[0]))):
        # Word-wrap the inline content in this block element.
        # First, find out if this is an li element with a trailing list.
        if ht.name == 'li' and isinstance(ht.content[-1], Element) and ht.content[-1].name in ('ol', 'ul', 'table'):
            content_to_wrap = content[:-1]
            last_block = content[-1]
        else:
            content_to_wrap = content
            last_block = None

        # Dump content_to_wrap to a temporary buffer.
        tmpf = StringIO()
        tmpf.write(start_tag(ht))
        write_inline_content(tmpf, content_to_wrap, indent + '  ', allow_last_block=False, strict=strict)
        if last_block is None:
            tmpf.write('</{}>'.format(ht.name))
        text = tmpf.getvalue()

        # Write the output to f.
        if '\n' in text:
            # This is unexpected; don't word-wrap. Write it verbatim, with newlines.
            f.write(indent + text + "\n")
        else:
            # The usual case. Word-wrap and write.
            subsequent_indent = indent
            if ht.name != 'p':
                subsequent_indent += '    '
            f.write(textwrap.fill(text, WIDTH, expand_tabs=False, replace_whitespace=False,
                                  initial_indent=indent, subsequent_indent=subsequent_indent,
                                  break_long_words=False, break_on_hyphens=False))
            f.write("\n")

        # If we had a trailing block, dump it now (and the end tag we skipped before).
        if last_block:
            write_html(f, last_block, indent + '  ', strict=strict)
            f.write(indent + "</{}>\n".format(ht.name))

    elif info[0] == 'B':
        # Block.
        f.write(indent + start_tag(ht))
        if info != 'B0':
            if content:
                if is_ht_inline(content[0]):
                    if strict and info[1] not in 'i?s':
                        if isinstance(content[0], str):
                            raise ValueError("<{}> element may only contain tags, not text".format(ht.name))
                        else:
                            raise ValueError("<{}> element may not contain inline element <{}>".format(ht.name, content[0].name))
                    write_inline_content(f, content, indent + '  ', ht.name == 'li', strict)
                else:
                    if strict and info[1] not in 'b?':
                        raise ValueError("<{}> element may not contain block element <{}>".format(ht.name, content[0].name))
                    inner_indent = indent
                    if ht.name not in non_indenting_tags:
                        inner_indent += '  '
                    f.write('\n')
                    prev_needs_space = False
                    first = True
                    for kid in content:
                        if strict and is_ht_inline(kid):
                            if isinstance(kid, str):
                                raise ValueError("<{}> element may contain either text or block content, not both".format(ht.name))
                            else:
                                raise ValueError("<{}> element may contain either blocks (like <{}>) "
                                                 "or inline content (like <{}>), not both".format(
                                        ht.name, content[0].name, kid.name))
                        needs_space = ((strict or isinstance(kid, Element))
                                       and (kid.name in {'p', 'figure', 'section'}
                                            or (kid.name in {'div', 'li', 'td'}
                                                and kid.content and not is_ht_inline(kid.content[0]))))
                        if not first and (prev_needs_space or needs_space):
                            f.write('\n')
                        write_html(f, kid, inner_indent, strict)
                        prev_needs_space = needs_space
                        first = False
                    f.write(indent)
            f.write('</{}>'.format(ht.name))
        f.write('\n')
    else:
        # Inline. Content must be inline too.
        assert info in ('Ii', 'I0')
        f.write(start_tag(ht))
        if info != 'I0':
            write_inline_content(f, content, indent + '  ', False, strict)
            f.write('</{}>'.format(ht.name))

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

    _tag_raw_data = '''\
    html head body section hgroup table tbody thead tfoot tr ol\
        ul blockquote ol ul dl figure: Bb
    title p h1 h2 h3 h4 h5 h6 address figcaption pre: Bi
    a em strong small s cite q dfn abbr data time code var samp\
        kbd sub sup i b u mark bdi bdo span: Ii
    br wbr: I0
    div li td th noscript object dt dd: B?
    meta link hr img: B0
    style script: Bs'''

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
