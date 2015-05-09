#!/bin/bash
# Upload the latest snapshot to https://people.mozilla.org/~jorendorff/es6-draft.html

set -eu
html-minifier --collapse-whitespace --remove-attribute-quotes \
    --remove-redundant-attributes --prevent-attributes-escaping \
    --use-short-doctype --remove-optional-tags -o es6-draft.html \
    es6-draft.html
sftp -b ftpcommands.txt people.mozilla.org

