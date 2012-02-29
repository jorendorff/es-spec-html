import html

## def all_elements(e):
##     yield e
##     for k in doc.contents:
##         if not isinstance(k, str):
##             for d in all_elements(k):
##                 yield d

def fixup_lists(e):
    if e.name in ('ol', 'ul'):
        # This is already a list. I don't think there are any lists we don't
        # identify during transform phase that are nested in lists we do
        # identify. So skip this.
        return

    kids = e.content

    have_list_items = False
    for k in kids:
        if not isinstance(k, str):
            fixup_lists(k)
            if k.name == 'li':
                have_list_items = True

    

    # Walk the elements from left to right. If we find any <li> elements,
    # wrap them in <ol> elements to the appropriate depth.
    new_content = []
    lists = []
    for k in kids:
        if isinstance(k, str) or k.name != 'li':
            # Not a list item. Close all open lists. Add k to new_content.
            del lists[:]
            new_content.append(k)
        else:
            # Oh no. It is a list item. Well, what is its depth?
            if k.style and '@num' in k.style:
                depth = int(k.style['@num'].partition('/')[0])

                # While we're here, delete the @num magic style attribute.
                del k.style['@num']
            else:
                depth = 0

            # Close any open lists at greater depth.
            while lists and lists[-1][0] > depth:
                del lists[-1]

            # If we don't already have a list at that depth, open one.
            if not lists or depth > lists[-1][0]:
                new_list = html.ol(class_='block' if lists else 'proc')

                # If there is an enclosing list, add new_list to the last <li>
                # of the enclosing list, not the enclosing list itself.
                # If there is no enclosing list, add new_list to new_content.
                (lists[-1][1][-1].content if lists else new_content).append(new_list)
                lists.append((depth, new_list.content))

            lists[-1][1].append(k)

    kids[:] = new_content

def fixup(doc):
    fixup_lists(doc)
    return doc
