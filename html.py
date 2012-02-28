from collections import namedtuple
from cgi import escape

Element = namedtuple('Element', 'name, attrs, content')

empty_tags = {'meta', 'br', 'hr'}

def save_html(filename, ht):
    assert ht.name == 'html'

    def htmlify(name):
        """ Convert a pythonified tag name or attribute name back to HTML. """
        if name.endswith('_'):
            name = name[:-1]
        name = name.replace('_', '-')
        return name

    def start_tag(ht):
        attrs = ''.join(' {0}="{1}"'.format(htmlify(k), escape(v, True))
                        for k, v in ht.attrs.items())
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
                    for k in ht.content:
                        write_block(f, k, indent + "  ")
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

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("<!doctype html>\n")
        write_block(f, ht)

__all__ = []  # modified by _init

def _init(v):
    def element_constructor(name):
        def construct(*content, **attrs):
            return Element(name, attrs, content)
        construct.__name__ = name
        return construct

    names = ('html head meta link '
             'body section div figure '
             'table caption colgroup col tbody thead tfoot tr td th '
             'h1 h2 p a b em i span strong sub sup br hr').split()

    for name in names:
        v[name] = element_constructor(name)
        __all__.append(name)

_init(vars())
