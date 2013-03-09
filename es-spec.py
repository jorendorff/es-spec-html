#!/usr/bin/env python3

import sys, os
import docx
import htmodel
from transform import transform
import fixups

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

def save_html(docxfile, filename):
    result = transform(docxfile.document)
    result = fixups.fixup(docxfile, result)

    print("=== Unrecognized styles")
    for k, v in sorted(fixups.unrecognized_styles.items(), key=lambda pair: pair[1]):
        print(k, v)
    print()

    htmodel.save_html(filename, result)

if len(sys.argv) != 2:
    print("usage: {} filename.docx".format(sys.argv[0]), file=sys.stderr)
    sys.exit(1)

in_filename = sys.argv[1]
base, ext = os.path.splitext(in_filename)
assert ext in ('.docx', '.dotx')
out_filename = base + '.html'
doc = docx.load(in_filename)
save_html(doc, out_filename)

# Some other things that can be done with a docx.Document:
#sketch_schema(doc.document)
#doc._extract()
#doc._dump_styles()

