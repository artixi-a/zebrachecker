import json
import os

PRESET_FILE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "presets.json"
)

BLANK = {
    "name": "Auto-detect (no preset)",
    "prefix": "",
    "gs_mode": "",
    "marker": "",
    "gs_prefix": "",
    "expected_length": 0,
    "dedupe_mode": "core_id",
}

FIELDS = [
    "prefix",
    "gs_mode",
    "marker",
    "gs_prefix",
    "expected_length",
    "dedupe_mode",
    "key_mode",
    "min_length",
]


def load_presets(path=None):
    path = path or PRESET_FILE
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        presets = data.get("presets") if isinstance(data, dict) else data
        if not presets:
            return [dict(BLANK)]
        return [dict(BLANK)] + [p for p in presets if p.get("name") != BLANK["name"]]
    except Exception:
        return [dict(BLANK)]


def names(presets):
    return [p.get("name", "(unnamed)") for p in presets]


def get(presets, name):
    for preset in presets:
        if preset.get("name") == name:
            return dict(preset)
    return dict(BLANK)


def parse_presets(raw):
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="ignore")
    data = json.loads(raw)
    if isinstance(data, dict):
        presets = data.get("presets", [])
    elif isinstance(data, list):
        presets = data
    else:
        presets = []
    return [p for p in presets if isinstance(p, dict) and p.get("name")]


def to_json_bytes(presets):
    cleaned = [p for p in presets if p.get("name") != BLANK["name"]]
    return json.dumps({"presets": cleaned}, indent=2).encode("utf-8")


def upsert(presets, preset):
    name = preset.get("name")
    if not name:
        return presets
    out = [p for p in presets if p.get("name") not in (name, BLANK["name"])]
    out.append(preset)
    return [dict(BLANK)] + out
