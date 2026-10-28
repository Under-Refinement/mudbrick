"""Word equations (OMML) for docx2pdf.

Every equation gets a stable key (a hash of its content). Simple inline math is turned
into ordinary text runs (italic variables, real sub/superscripts) so it flows and
justifies with the sentence; display equations and inline math too complex for text
(fractions, accents, matrices, sums, nested scripts...) become placeholders that
layout.jsx replaces with a rendered PDF from the equations folder (see render_math.py).
Only the Python standard library is used.
"""
import hashlib
import re

MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M = "{%s}" % MNS
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TOKEN = "@@MATH:%s@@"
TOKEN_RE = re.compile(r"@@MATH:([0-9a-f]{12})@@")

GREEK_LOWER = set("αβγδεζηθικλμνξοπρστυφχψωϵϑϕϖϱς")
RELATIONS = set("=<>≤≥≠≈≡∈∉⊂⊆←→↔⇒⇔∝∼≅")
BINARY = set("+−×±∓·⋅∗∘")
SCRIPT = {"l": "ℓ"}
DOUBLE = {"R": "ℝ", "N": "ℕ", "Z": "ℤ", "Q": "ℚ", "C": "ℂ"}
THIN = " "


class Complex(Exception):
    """This equation cannot be set as plain text."""


def local(el):
    return el.tag.split("}", 1)[-1] if "}" in el.tag else el.tag


def is_math(el):
    return el.tag in (M + "oMath", M + "oMathPara")


def canon(el):
    """Content-only string for hashing (ignores namespace prefixes and run ids)."""
    if el.tag.startswith(W) and local(el) in ("rPr",):
        return ""
    attrs = ",".join("%s=%s" % (k.split("}")[-1], v) for k, v in sorted(el.attrib.items())
                     if not k.split("}")[-1].startswith("rsid"))
    inner = "".join(canon(c) for c in el)
    return "%s[%s](%s%s)" % (local(el), attrs, el.text or "" if local(el) == "t" else "", inner)


def key(el):
    return hashlib.sha1(canon(el).encode("utf-8")).hexdigest()[:12]


def plain_text(el):
    return "".join(t.text or "" for t in el.iter(M + "t"))


# ---------------------------------------------------------------------------
# OMML -> text runs
# ---------------------------------------------------------------------------

def _val(pr, tag):
    if pr is None:
        return None
    x = pr.find(M + tag)
    return None if x is None else x.get(M + "val", "on")


def _run(r, script, upright=False):
    pr = r.find(M + "rPr")
    sty, scr, nor = _val(pr, "sty"), _val(pr, "scr"), pr is not None and pr.find(M + "nor") is not None
    text = "".join(t.text or "" for t in r.iter(M + "t"))
    # script/double-struck letters without a text form (e.g. script L) are rendered in
    # LaTeX, so they look the same inline as in display equations
    if scr == "script":
        if any(c.isalpha() and c not in SCRIPT for c in text):
            raise Complex()
        text = "".join(SCRIPT.get(c, c) for c in text)
    elif scr == "double-struck":
        if any(c.isalpha() and c not in DOUBLE for c in text):
            raise Complex()
        text = "".join(DOUBLE.get(c, c) for c in text)
    elif scr in ("fraktur",):
        raise Complex()
    text = text.replace("-", "−")
    bold = sty in ("b", "bi")
    out = []
    for c in text:
        if nor or upright or sty in ("p", "b") or scr == "sans-serif":
            ital = False
        elif sty in ("i", "bi"):
            ital = c.isalpha()
        else:
            ital = ("a" <= c.lower() <= "z") or c in GREEK_LOWER
        f = {"b": bold, "i": ital, "sup": script == "sup", "sub": script == "sub",
             "u": False, "mono": False}
        if out and out[-1][1] == f:
            out[-1] = (out[-1][0] + c, f)
        else:
            out.append((c, f))
    return out


def _children(el, script, upright=False):
    out = []
    for c in el:
        out += _node(c, script, upright)
    return out


def _node(el, script, upright=False):
    n = local(el)
    if not el.tag.startswith(M):
        return []                               # w:rPr, bookmarks, ...
    if n.endswith("Pr") or n in ("ctrlPr",):
        return []
    if n == "r":
        return _run(el, script, upright)
    if n in ("oMath", "e", "box", "num", "den", "fName", "lim", "deg"):
        return _children(el, script, upright)
    if n in ("sSub", "sSup"):
        if script:
            raise Complex()                     # nested scripts don't fit in text
        base = _children(el.find(M + "e"), script, upright)
        part = el.find(M + ("sub" if n == "sSub" else "sup"))
        return base + _children(part, "sub" if n == "sSub" else "sup", upright)
    if n == "d":
        pr = el.find(M + "dPr")
        beg = _val(pr, "begChr")
        end = _val(pr, "endChr")
        sep = _val(pr, "sepChr")
        beg = "(" if beg is None else ("" if beg == "on" else beg)
        end = ")" if end is None else ("" if end == "on" else end)
        sep = "|" if sep is None else sep
        parts = [e for e in el if local(e) == "e"]
        f = {"b": False, "i": False, "sup": script == "sup", "sub": script == "sub",
             "u": False, "mono": False}
        out = [(beg, f)] if beg else []
        for k, e in enumerate(parts):
            if k:
                out.append((sep, f))
            out += _children(e, script, upright)
        return out + ([(end, f)] if end else [])
    if n == "func":
        name = _children(el.find(M + "fName"), script, upright=True)
        arg = _children(el.find(M + "e"), script, upright)
        joiner = [] if arg and arg[0][0][:1] in "([{" else [(THIN, dict(name[-1][1]))] if name else []
        return name + joiner + arg
    raise Complex()                             # f, acc, bar, nary, m, eqArr, limLow, rad, sSubSup, ...


def _space_operators(runs):
    """Thin spaces around relations and binary operators (not in scripts, not unary)."""
    out, prev = [], ""
    for t, f in runs:
        if f["sub"] or f["sup"]:
            out.append((t, f)); prev = t[-1:] or prev
            continue
        s = ""
        for c in t:
            if c in RELATIONS or (c in BINARY and prev and prev not in "([{" and prev not in RELATIONS
                                   and prev not in BINARY and prev != " "):
                s = s.rstrip(" ") + THIN + c + THIN
            elif c == " " and s.endswith(THIN):
                continue
            else:
                s += c
            prev = c
        out.append((s, f))
    return out


def to_runs(el):
    """Text runs for a simple equation, or None if it needs rendering."""
    try:
        runs = _space_operators(_node(el if local(el) == "oMath" else el.find(M + "oMath"), None))
    except Complex:
        return None
    if not any(t.strip() for t, _ in runs):
        return None
    return runs
