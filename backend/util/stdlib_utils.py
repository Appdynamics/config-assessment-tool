import base64
import binascii
from enum import Enum
from typing import Optional


def substringBetween(s: str, left: str, right: str) -> Optional[str]:
    """Returns substring between two chars. Returns"""
    try:
        return (s.split(left))[1].split(right)[0]
    except IndexError:
        return None


def get_recursively(search_dict, field):
    """
    Takes a dict with nested lists and dicts, and searches all dicts for a key of the field provided.
    """
    fields_found = set()

    for key, value in search_dict.items():
        if key == field:
            if isinstance(value, list):
                for x in value:
                    fields_found.add(x)
            else:
                fields_found.add(value)
        elif isinstance(value, dict):
            results = get_recursively(value, field)
            for result in results:
                fields_found.add(result)
        elif isinstance(value, list):
            for item in value:
                if isinstance(item, dict):
                    more_results = get_recursively(item, field)
                    for another_result in more_results:
                        fields_found.add(another_result)
    return fields_found


def jsonEncoder(o):
    if isinstance(o, set):
        return list(o)
    if isinstance(o, Enum):
        return o.name
    if hasattr(o, "__json__"):
        return o.__json__()
    else:
        return f"<<non-serializable: {type(o).__qualname__}>>"


def isBase64(s: str, encoding="ISO-8859-1"):
    try:
        strBytes = bytes(s, encoding=encoding)
        return base64.b64encode(base64.b64decode(strBytes)) == strBytes
    except binascii.Error:
        return False


def base64Encode(s: str, encoding="ISO-8859-1"):
    strBytes = bytes(s, encoding=encoding)
    encodedBytes = base64.standard_b64encode(strBytes)
    return str(encodedBytes)[2:-1]


def base64Decode(s: str, encoding="ISO-8859-1"):
    strBytes = bytes(s, encoding=encoding)
    decodedBytes = base64.standard_b64decode(strBytes)
    return str(decodedBytes)[2:-1]
