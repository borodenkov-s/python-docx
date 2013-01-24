#-*- coding:utf-8 -*-

def findTypeParent(element, tag):
    """ Finds fist parent of element of the given type
    @param object element: etree element
    @param string the tag parent to search for

    @return object element: the found parent or None when not found
    """

    p = element
    while True:
        p = p.getparent()
        if p.tag == tag:
            return p

    # Not found
    return None
