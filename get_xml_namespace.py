"""Function to parse xml for namespace, dynamically."""

from xml.etree import cElementTree as et
from io import StringIO


def ucd(input):
    """Convert a string or unicode object, to a unicode version."""
    # print(type(input), input)
    try:
        input = input.decode('utf-8')
    except AttributeError:
        pass
    return input


def getns(xml):
    """Get a dictionary of namespaces from string or unicode."""
    # Thanks to Davide Brunato on stackoverflow for his answer on this bit.
    # https://stackoverflow.com/a/37409050/10941169
    xml = ucd(xml)
    namespaces = dict([
        node for _, node in et.iterparse(
            StringIO(xml), events=['start-ns']
            # StringIO(str(xml), events=['start-ns']
        )
    ])
    return namespaces