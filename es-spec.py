#!/usr/bin/env python3.2

import zipfile
from xml.etree import ElementTree
from transform import transform, shorten
import html
from cgi import escape

with zipfile.ZipFile("es6-draft.docx") as f:
    document = ElementTree.fromstring(f.read('word/document.xml'))

## [body] = list(document)
## for e in body:
##     if e.tag != u'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
##         print(e.tag)

def sketch_schema(document):
    def has_own_text(e):
        return e.text or any(k.tail for k in e)

    hits = set()
    def walk(e, path):
        tag = shorten(e.tag)
        path.append(tag)
        hits.add(' > '.join(path))
        if tag != 'pict':
            kpath = path
            if tag in ('r', 'p') or tag.endswith('Pr') or tag.endswith('PrEx'):
                kpath = [tag]
            for kid in e:
                walk(kid, kpath)
        path.pop()

    walk(document, [])

    for path in sorted(hits):
        print(path)

    print(len(hits))

#sketch_schema(document)

def save_xml(document):
    def writexml(e, out, indent='', context='block'):
        t = shorten(e.tag)
        assert e.tail is None
        start_tag = t
        for k, v in e.items():
            start_tag += ' {0}="{1}"'.format(shorten(k), escape(v, True))

        kids = list(e)
        if kids:
            assert e.text is None
            out.write("{0}<{1}>\n".format(indent, start_tag))
            for k in kids:
                writexml(k, out, indent + '  ')
            out.write("{0}</{1}>\n".format(indent, t))
        elif e.text:
            out.write("{0}<{1}>{2}</{3}>\n".format(indent, start_tag, escape(e.text), t))
        else:
            out.write("{0}<{1} />\n".format(indent, start_tag))

    with open('original.xml', 'w', encoding='utf-8') as out:
        writexml(document, out)
        out.write("\n")

#save_xml(document)

def save_html(document):
    result = transform(document)
    html.save_html('doc.html', result)

save_html(document)
