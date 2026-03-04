from gender_guesser.detector import Detector
import re

_DET = Detector(case_sensitive=False)

_PREFIXES = {"mr","mrs","ms","miss","mx","dr","prof","sir","madam","lord","lady"}
_SUFFIXES = {"jr","sr","ii","iii","iv","phd","md","esq"}

def _tokenize_first_name(raw: str) -> str:
    if raw is None:
        return ""
    s = str(raw).strip()

    if "," in s and s.count(",") == 1:
        left, right = [t.strip() for t in s.split(",", 1)]
        if len(right.split()) <= 3:
            s = right

    s = re.sub(r"\(.*?\)|\[.*?\]|\".*?\"", "", s).strip()

    toks = re.split(r"[^\w\-’']+", s, flags=re.UNICODE)
    toks = [t for t in toks if t]

    if toks and toks[0].lower().strip(".") in _PREFIXES:
        toks = toks[1:]
    if toks and toks[-1].lower().strip(".") in _SUFFIXES:
        toks = toks[:-1]

    if not toks:
        return ""

    first = toks[0]

    if re.fullmatch(r"[A-Za-z]\.?$", first):
        return ""

    return first

def classify_first_name(raw: str) -> str:
    first = _tokenize_first_name(raw)
    if not first:
        return "unknown"

    g = _DET.get_gender(first)

    mapping = {
        "male": "m",
        "mostly_male": "m",
        "female": "f",
        "mostly_female": "f",
        "andy": "unknown",
        "unknown": "unknown"
    }
    out = mapping.get(g, "unknown")

    if out == "unknown" and "-" in first:
        g2 = _DET.get_gender(first.split("-")[0])
        out = mapping.get(g2, "unknown")

    return out
