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

    def find(self, matcher):
        match_result = matcher(self)
        assert match_result is True or match_result is False or match_result is None
        if match_result is None:
            return
        for kid in self.content:
            if isinstance(kid, Element):
                for x in kid.find(matcher):
                    yield x
        if match_result:
            yield self

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
        assert isinstance(name, str) # detect common bug
        return self.find_replace(lambda e: e.name == name, replacement)

def escape(s, quote=False):
    """ Escape the string s for HTML output.

    This escapes characters that are special in HTML (& < >) and all non-ASCII characters.
    If 'quote' is true, escape quotes (' ") as well.

    Why use character entity references for non-ASCII characters? The program
    encodes the output as UTF-8, so we should be fine without escaping. We
    escape only for maximum robustness against broken tools.
    """

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

def save_html(filename, ht, strict=True):
    assert ht.name == 'html'

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("<!doctype html>\n")
        write_html(f, ht, strict=strict)

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
                        for k, v in sorted(ht.attrs.items()))
        if ht.style:
            assert 'style' not in ht.attrs
            style = '; '.join(name + ": " + value for name, value in sorted(ht.style.items()))
            attrs += ' style="{0}"'.format(style)
        return '<{0}{1}>'.format(ht.name, attrs)

    def is_ht_inline(ht):
        return isinstance(ht, str) or _tag_data[ht.name][0] == 'I'

    def write_inline_content(f, content, indent, allow_last_block, strict, strict_blame):
        if (allow_last_block
              and isinstance(content[-1], Element)
              and content[-1].name in ('ol', 'ul', 'table', 'figure')):
            last = content[-1]
            content = content[:-1]
        else:
            last = None

        for kid in content:
            if isinstance(kid, str):
                f.write(escape(kid))
            else:
                if strict and not is_ht_inline(kid):
                    raise ValueError("block element <{}> can't appear in inline content:\n".format(kid.name)
                                     + repr(strict_blame))
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
        write_inline_content(tmpf, content_to_wrap, indent + '  ',
                             allow_last_block=ht.name == 'li', strict=strict, strict_blame=ht)
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
                    write_inline_content(f, content, indent + '  ', ht.name == 'li', strict, strict_blame=ht)
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
            write_inline_content(f, content, indent + '  ', False, strict, ht)
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

__all__ = ['Element'] + list(_tag_data.keys())
