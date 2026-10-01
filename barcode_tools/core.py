import re

GS_RAW = "\x1d"
GS_XML = "_x001D_"
GS_ALT = "!"

DEFAULT_MARKER = "91EE12"
DEFAULT_GS_PREFIX = "93"
DEFAULT_CORE_LENGTH = 31
DEFAULT_MIN_LENGTH = 80

COMMON_PREFIXES = [
    "0108606018940011215",
    "0104680679601959215",
    "0108606018945290215",
]

_XML_GS_RE = re.compile(r"_x001[dD]_")
_GS_PRESENCE_RE = re.compile(r"(?:_x001[dD]_|\x1d)")
_AIM_RE = re.compile(r"^\]d2")
_CONTROL_QUOTES_RE = re.compile(r'[\s\x00-\x1C\x1E-\x1F\x7F"]+')
_CONTROL_PLAIN_RE = re.compile(r"[\s\x00-\x1F\x7F]+")
_CONTROL_KEEP_GS_RE = re.compile(r'[\s\x00-\x1C\x1E-\x1F\x7F"]+')


def unxml(value):
    if value is None:
        return ""
    return _XML_GS_RE.sub(GS_RAW, str(value))


def has_gs(value):
    if value is None:
        return False
    return bool(_GS_PRESENCE_RE.search(str(value)))


def has_raw_gs(value):
    return value is not None and GS_RAW in str(value)


def has_xml_gs(value):
    return value is not None and bool(_XML_GS_RE.search(str(value)))


def gs_count(value):
    if value is None:
        return 0
    return len(_XML_GS_RE.findall(str(value))) + str(value).count(GS_RAW)


def strip_gs(value):
    if value is None:
        return ""
    return _XML_GS_RE.sub("", str(value)).replace(GS_RAW, "")


def strip_aim(value):
    if value is None:
        return ""
    return _AIM_RE.sub("", str(value))


def strip_controls(value, drop_quotes=True):
    if value is None:
        return ""
    pattern = _CONTROL_QUOTES_RE if drop_quotes else _CONTROL_PLAIN_RE
    return pattern.sub("", str(value))


def normalize(
    value,
    keep_gs=False,
    upper=True,
    strip_aim_prefix=True,
    drop_quotes=True,
):
    s = str(value) if value is not None else ""
    if not keep_gs:
        s = _XML_GS_RE.sub("", s).replace(GS_RAW, "")
    if keep_gs:
        s = _CONTROL_KEEP_GS_RE.sub("", s)
    else:
        s = strip_controls(s, drop_quotes=drop_quotes)
    if strip_aim_prefix:
        s = _AIM_RE.sub("", s)
    if upper:
        s = s.upper()
    return s


def light_clean(value, upper=True):
    s = str(value) if value is not None else ""
    s = _XML_GS_RE.sub("", s).replace(GS_RAW, "")
    s = s.strip()
    if upper:
        s = s.upper()
    return s


def core_id(value, length=DEFAULT_CORE_LENGTH, keep_gs=False, upper=True):
    return normalize(value, keep_gs=keep_gs, upper=upper)[:length]


def match_key(value, key_mode="core_id", length=DEFAULT_CORE_LENGTH):
    if key_mode == "full":
        return light_clean(value, upper=True)
    return core_id(value, length=length)


def gs_regex(target, alt=True):
    tokens = [r"_x001[dD]_", r"\x1d"]
    if alt:
        tokens.append(r"!")
    return re.compile("(?:" + "|".join(tokens) + ")" + re.escape(target))


def has_gs_before(value, target, alt=True):
    if value is None:
        return False
    return bool(gs_regex(target, alt=alt).search(str(value)))


def has_gs_surrounding(value, marker):
    if value is None:
        return False
    pattern = re.compile(
        r"(?:_x001[dD]_|\x1d)" + re.escape(marker) + r"(?:_x001[dD]_|\x1d)"
    )
    return bool(pattern.search(str(value)))


_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_CONTROL_VIS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")


def to_excel_safe(value):
    if not isinstance(value, str):
        return value
    s = value.replace(GS_RAW, GS_XML)
    return _ILLEGAL_XML_RE.sub(
        lambda m: "_x{:04X}_".format(ord(m.group(0))), s
    )


def has_invisible(value):
    if not isinstance(value, str):
        return False
    return has_gs(value) or bool(_CONTROL_VIS_RE.search(value))


def visualize(value, gs_token="[GS]"):
    if not isinstance(value, str):
        return value
    s = _XML_GS_RE.sub(gs_token, value).replace(GS_RAW, gs_token)
    return _CONTROL_VIS_RE.sub(
        lambda m: "[\\x{:02X}]".format(ord(m.group(0))), s
    )
