"""hacks.py - Mark the bad hacks in your code.

It's a fact of life that bad hacks are sometimes useful. The crufty special
case. The bizarre data tweak to cope with badness from upstream. However we
don't want them cluttering up our code after they are no longer being used.

Use declare_hack(str) to announce that there is a hack in your code.

Use using_hack(str) with the same string to note that the hack is still used.

At the end of your program, call warn_about_unused_hacks() to warn about every
hack that is declared but not used. Those are the ones you can delete.

"""

import warnings

_hacks = {}

def declare_hack(name):
    if name not in _hacks:
        _hacks[name] = False

def using_hack(name):
    assert name in _hacks
    _hacks[name] = True

def warn_about_unused_hacks():
    for name in sorted(_hacks):
        if _hacks[name] == False:
            warnings.warn("Hack %r is not used." % name)
