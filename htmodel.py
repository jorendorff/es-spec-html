import re
import html as _html
from html.entities import codepoint2name as _entities
from io import StringIO

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

    def to_html(self):
        f = StringIO()
        write_html(f, self)
        return f.getvalue()

    def kids(self, name=None):
        """ Return an iterator over (index, child) pairs where each child is a
        child element of self. The iterator is robust against mutating
        self.content to the right of i. """
        for i, kid in enumerate(self.content):
            if isinstance(kid, Element) and (name is None or kid.name == name):
                yield i, kid

# These tags all insist on being emitted on a line (or more) of their own.
# These tags all have in common that inserting space before and/or after them
# does not affect rendering.
_spaceable_tags = set('html head title base link meta style '
                      'table caption colgroup col tbody thead tfoot tr td th '
                      'body section nav article aside h1 h2 h3 h4 h5 h6 header footer '
                      'p hr pre blockquote ol ul li dl dt dd figure figcaption object div '.split())

def is_spaceable(ht):
    return not isinstance(ht, str) and ht.name in _spaceable_tags

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
non_indenting_tags = {'html', 'body', 'section'}

def save_html(filename, ht):
    assert ht.name == 'html'

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("<!doctype html>\n")
        write_html(f, ht)

def write_html(f, ht):
    def htmlify(name):
        """ Convert a pythonified tag name or attribute name back to HTML. """
        if name.endswith('_'):
            name = name[:-1]
        name = name.replace('_', '-')
        return name

    def start_tag(ht):
        attrs = ''.join(' {0}="{1}"'.format(htmlify(k), escape(v, True))
                        for k, v in ht.attrs.items())
        if ht.style and 'style' not in ht.attrs:
            style = '; '.join(name + ": " + value for name, value in sorted(ht.style.items()))
            attrs += ' style="{0}"'.format(style)
        return '<{0}{1}>'.format(ht.name, attrs)

    def write_block(f, ht, indent=''):
        if isinstance(ht, str):
            f.write(indent + escape(ht) + '\n')
        else:
            f.write(indent + start_tag(ht))
            if ht.name in empty_tags and not ht.content:
                f.write("\n")
            else:
                inner_indent = indent
                if ht.name not in non_indenting_tags:
                    inner_indent += "  "

                if any(not is_spaceable(k) for k in ht.content):
                    write_inline_content(f, ht.content, inner_indent)
                else:
                    f.write("\n")
                    for k in ht.content:
                        write_block(f, k, inner_indent)
                    f.write(indent)
                f.write("</{0}>\n".format(ht.name))

    def write_inline_content(f, content, indent):
        for k in content:
            if not isinstance(k, str) and k.name in ('ol', 'ul', 'table'):
                f.write('\n')
                write_block(f, k, indent)
                f.write(indent)
            else:
                write_inline(f, k, indent)

    def write_inline(f, ht, indent):
        if isinstance(ht, str):
            f.write(escape(ht))
        else:
            f.write(start_tag(ht))
            write_inline_content(f, ht.content, indent)
            if ht.content or ht.name not in empty_tags:
                f.write("</{0}>".format(ht.name))

    write_block(f, ht)

__all__ = []  # modified by _init

_seen = set()

def _init(v):
    def element_constructor(name):
        def construct(*content, **attrs):
            _seen.add(name)
            if 'class_' in attrs:
                attrs['class'] = attrs['class_']
                del attrs['class_']
            return Element(name, attrs, None, list(content))
        construct.__name__ = name
        return construct

    names = ('html head title base link meta style script noscript '
             'body section nav article aside h1 h2 h3 h4 h5 h6 hgroup header footer address '
             'p hr pre blockquote ol ul li dl dt dd figure figcaption div '
             'a em strong small s cite q dfn abbr data time code var '
             'samp kbd sub sup i b u mark ruby rt rp bdi bdo span br wbr '
             'img iframe embed object param video audio source track canvas map area '
             'table caption colgroup col tbody thead tfoot tr td th').split()

    for name in names:
        v[name] = element_constructor(name)
        __all__.append(name)

_init(vars())
