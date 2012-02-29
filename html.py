from collections import namedtuple
from cgi import escape
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

empty_tags = {'meta', 'br', 'hr'}
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
                if ht.name != 'body' and any(isinstance(k, str) for k in ht.content):
                    for k in ht.content:
                        write_inline(f, k)
                else:
                    f.write("\n")
                    inner_indent = indent
                    if ht.name not in non_indenting_tags:
                        inner_indent += "  "
                    for k in ht.content:
                        write_block(f, k, inner_indent)
                    f.write(indent)
                f.write("</{0}>\n".format(ht.name))

    def write_inline(f, ht):
        if isinstance(ht, str):
            f.write(escape(ht))
        else:
            f.write(start_tag(ht))
            if ht.content or ht.name not in empty_tags:
                for k in ht.content:
                    write_inline(f, k)
                f.write("</{0}>".format(ht.name))

    write_block(f, ht)

__all__ = []  # modified by _init

def _init(v):
    def element_constructor(name):
        def construct(*content, **attrs):
            return Element(name, attrs, None, list(content))
        construct.__name__ = name
        return construct

    names = ('html head meta link '
             'table caption colgroup col tbody thead tfoot tr td th '
             'body section nav article aside h1 h2 h3 h4 h5 h6 hgroup header footer address '
             'p hr pre blockquote ol ul li dl dt dd figure figcaption div '
             'a em strong small s cite q dfn abbr data time code var '
             'samp kbd sub sup i b u mark ruby rt rp bdi bdo span br wbr').split()

    for name in names:
        v[name] = element_constructor(name)
        __all__.append(name)

_init(vars())
