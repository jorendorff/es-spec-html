# es-spec - Convert the ECMAScript Language Specification to HTML

To install the HTML minifier dependency using [npm](https://www.npmjs.com/):

    npm install html-minifier -g

To run the program:

    ./es-spec.py es6-draft.docx

Note: Python 3 is required.

To minify the generated HTML afterwards:

    html-minifier --collapse-whitespace --remove-attribute-quotes \
        --remove-redundant-attributes --prevent-attributes-escaping \
        --use-short-doctype --remove-optional-tags -o es6-draft.html \
        es6-draft.html

## About this program

**Architecture:** The program is in four parts:

  * Load the Word document (`docx.py`)
  * Convert it to extremely rough HTML+CSS (`transform.py`)
  * Apply a series of transformations, ranging from minor tweaks to
    very fancy algorithms, to the HTML (`fixups.py`)
  * Dump the resulting HTML document (`htmodel.py`)

Most of the interesting work, and most of the bugs, are in `fixups.py`.

**Fragility:** The script is quite sensitive to the input document and
will throw an exception and give up if the document isn't exactly as
expected.  It's been hard to balance (a) being "liberal in what you
accept" with (b) making sure fixups do not break silently, but rather
get the user's attention, when the input document changes in unexpected
ways.

**Debugging:** If a directory named `_fixup_log` exists under the
current directory, the script dumps the whole halfway-transformed
document to a file in that directory after each fixup.
