#!/usr/bin/env python3.2

import docx
import htmodel
from transform import transform, unrecognized_styles
import fixups

es6_draft = docx.load("es6-draft.docx")

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

def save_html(document):
    result = transform(document)

    print("=== Unrecognized styles")
    for k, v in sorted(unrecognized_styles.items(), key=lambda pair: pair[1]):
        print(k, v)
    print()

    fixups.fixup(result)
    htmodel.save_html('es6-draft.html', result)

#sketch_schema(es6_draft.document)
#es6_draft._extract()
save_html(es6_draft.document)
