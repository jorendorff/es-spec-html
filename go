#!/bin/bash
# Upload the latest snapshot to http://people.mozilla.org/~jorendorff/es6-draft.html

set -eu
sftp -b ftpcommands.txt people.mozilla.org

