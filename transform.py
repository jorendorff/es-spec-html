from html import *

namespaces = {
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main': '',
    'http://schemas.openxmlformats.org/markup-compatibility/2006': 'compat',
    'urn:schemas-microsoft-com:vml': 'vml',
    'urn:schemas-microsoft-com:office:office': 'office',
    'http://www.w3.org/XML/1998/namespace': 'xml',
    'urn:schemas-microsoft-com:office:word': 'msword',
}

def shorten(name):
    if name[:1] == '{':
        end = name.index('}')
        schema = name[1:end]
        v = namespaces.get(schema)
        if v is None:
            return name
        elif v == '':
            return name[end + 1:]
        else:
            return v + ':' + name[end + 1:]
    else:
        return name

def transform(e):
    name = shorten(e.tag)
    assert e.tail is None

    if name == 't':
        assert len(e) == 0
        return e.text
    elif name == 'instrText':
        assert len(e) == 0
        return '{' + e.text + '}'
    else:
        assert e.text is None

        # Transform all children.
        c = []
        for k in e:
            ht = transform(k)
            if ht is not None:
                c.append(ht)

        if name == 'document':
            [body_e] = c
            return html(
                head(
                    meta(http_equiv="Content-Type", content="text/html; charset=UTF-8"),
                    link(rel="stylesheet", type="text/css", href="es-spec.css")),
                body_e)
        elif name == 'body':
            return body(section(*c, id="everything"))
        elif name == 'p':
            return p(*c)
        elif name == 'br':
            assert not c
            assert set(e.keys()) <= {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type'}
            br_type = e.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
            if br_type is None:
                return br()
            else:
                assert br_type == 'page'
                return hr()
        else:
            if len(c) == 0:
                return None
            elif len(c) == 1:
                return c[0]
            else:
                return div(*c)

__all__ = ['transform', 'shorten']
