import base64


def base64Encode(s: str, encoding="ISO-8859-1"):
    strBytes = bytes(s, encoding=encoding)
    encodedBytes = base64.standard_b64encode(strBytes)
    return str(encodedBytes)[2:-1]
