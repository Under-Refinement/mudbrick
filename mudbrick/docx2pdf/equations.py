"""Word equations (OMML) -> LaTeX -> typeset PDF, set in the body font.

Stand-alone (display) equations are converted to LaTeX and typeset by a TeX engine
(Tectonic, XeLaTeX or LuaLaTeX, whichever is installed) with unicode-math: letters,
digits, Greek and + - = ( ) come from the body font (Myriad Pro), and only the symbols
it lacks (big sums, integrals, fraction rules, ...) from a sans math font. Each equation
becomes one page of equations.pdf, cropped exactly to its box, so InDesign can place it
inline on the text baseline. Without a TeX engine the equations stay as plain text.
"""
import os
import re
import shutil
import subprocess

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

FUNCTIONS = {"log", "ln", "lg", "exp", "sin", "cos", "tan", "cot", "sec", "csc", "sinh",
             "cosh", "tanh", "coth", "arcsin", "arccos", "arctan", "arg", "min", "max",
             "lim", "sup", "inf", "det", "dim", "ker", "deg", "gcd", "Pr", "liminf",
             "limsup", "argmin", "argmax"}
NARY = {"\u2211": r"\sum", "\u220F": r"\prod", "\u2210": r"\coprod", "\u222B": r"\int",
        "\u222C": r"\iint", "\u222D": r"\iiint", "\u222E": r"\oint", "\u22C3": r"\bigcup",
        "\u22C2": r"\bigcap", "\u22C1": r"\bigvee", "\u22C0": r"\bigwedge"}
ACCENT = {"\u0302": r"\hat", "\u0303": r"\tilde", "\u0304": r"\bar", "\u0305": r"\bar",
          "\u0307": r"\dot", "\u0308": r"\ddot", "\u20D7": r"\vec", "\u0306": r"\breve",
          "\u030C": r"\check", "\u0301": r"\acute", "\u0300": r"\grave"}
DELIM = {"(": "(", ")": ")", "[": "[", "]": "]", "{": r"\{", "}": r"\}", "|": "|",
         "\u2016": r"\Vert", "\u27E8": r"\langle", "\u27E9": r"\rangle", "\u2329": r"\langle",
         "\u232A": r"\rangle", "\u230A": r"\lfloor", "\u230B": r"\rfloor", "\u2308": r"\lceil",
         "\u2309": r"\rceil", "": "."}
CHARS = {"\\": r"\backslash ", "{": r"\{", "}": r"\}", "#": r"\#", "%": r"\%", "&": r"\&",
         "$": r"\$", "_": r"\_", "^": r"\hat{}", "~": r"\sim ", "*": r"\ast ",
         "-": "\u2212", "\u2013": "\u2212"}   # hyphen -> the body font's true minus


def _val(el, tag, default=None):
    """m:val of the first `tag` under el's properties."""
    if el is None:
        return default
    x = el.find(".//" + M + tag)
    return default if x is None else x.get(M + "val", default)


def _text(s):
    """Escape TeX specials; words of 3+ letters (and known functions) set upright."""
    s = re.sub(r"\s+", "", s)
    out = []
    for tok in re.findall(r"[A-Za-z]+|.", s):
        if tok in FUNCTIONS:
            out.append("\\" + tok + " " if tok not in ("argmin", "argmax")
                       else r"\operatorname*{%s}" % tok[3:].join(("arg", "")))
        elif len(tok) >= 3 and tok.isalpha():
            out.append(r"\mathrm{%s}" % tok)
        else:
            out.append(CHARS.get(tok, tok))
    return "".join(out)


class Converter:
    def __init__(self, tracked="accept"):
        self.tracked = tracked

    def kids(self, el, tag=None):
        if el is None:
            return ""
        if tag is not None:
            el = el.find(M + tag)
            if el is None:
                return ""
        return "".join(self.conv(c) for c in el)

    def arg(self, el, tag):
        return "{" + self.kids(el, tag) + "}"

    def conv(self, el):
        t = el.tag
        name = t[len(M):] if t.startswith(M) else None
        if t in (W + "ins", W + "moveTo"):          # tracked insertion around math
            return "" if self.tracked == "reject" else self.kids(el)
        if t in (W + "del", W + "moveFrom"):
            return self.kids(el) if self.tracked == "reject" else ""
        if t == W + "r":                            # ordinary text inside an equation
            txt = "".join(x.text or "" for x in el.iter(W + "t"))
            return r"\text{%s}" % txt if txt.strip() else ""
        if name is None or name.endswith("Pr"):
            return ""
        if name == "r":
            if el.find(W + "del") is not None and self.tracked != "reject":
                return ""
            if el.find(W + "ins") is not None and self.tracked == "reject":
                return ""
            txt = "".join(x.text or "" for x in el.iter(M + "t"))
            if not txt:
                return ""
            rpr = el.find(M + "rPr")
            if rpr is not None and rpr.find(M + "nor") is not None:
                return r"\text{%s}" % txt
            sty = _val(rpr, "sty")
            body = _text(txt)
            if sty == "p":
                return r"\mathrm{%s}" % body
            if sty == "b":
                return r"\mathbf{%s}" % body
            if sty == "bi":
                return r"\symbfit{%s}" % body
            return body
        if name in ("oMath", "e", "num", "den", "sub", "sup", "deg", "fName", "lim",
                    "box", "borderBox", "phant"):
            return self.kids(el)
        if name == "oMathPara":
            return r" \\ ".join(self.conv(c) for c in el if c.tag == M + "oMath")
        if name == "sSub":
            return "{%s}_%s" % (self.kids(el, "e"), self.arg(el, "sub"))
        if name == "sSup":
            return "{%s}^%s" % (self.kids(el, "e"), self.arg(el, "sup"))
        if name == "sSubSup":
            return "{%s}_%s^%s" % (self.kids(el, "e"), self.arg(el, "sub"), self.arg(el, "sup"))
        if name == "sPre":
            return "{}_%s^%s{%s}" % (self.arg(el, "sub"), self.arg(el, "sup"), self.kids(el, "e"))
        if name == "f":
            num, den = self.kids(el, "num"), self.kids(el, "den")
            kind = _val(el.find(M + "fPr"), "type", "bar")
            if kind == "noBar":
                return r"\genfrac{}{}{0pt}{}{%s}{%s}" % (num, den)
            if kind in ("lin", "skw"):
                return "{%s}/{%s}" % (num, den)
            return r"\frac{%s}{%s}" % (num, den)
        if name == "nary":
            pr = el.find(M + "naryPr")
            op = NARY.get(_val(pr, "chr", "\u222B"), r"\int")
            loc = _val(pr, "limLoc")
            if loc == "undOvr":
                op += r"\limits"
            elif loc == "subSup" and op not in (r"\int", r"\iint", r"\iiint", r"\oint"):
                op += r"\nolimits"
            if _val(pr, "subHide") not in ("1", "on", "true"):
                sub = self.kids(el, "sub")
                if sub:
                    op += "_{%s}" % sub
            if _val(pr, "supHide") not in ("1", "on", "true"):
                sup = self.kids(el, "sup")
                if sup:
                    op += "^{%s}" % sup
            return op + "{" + self.kids(el, "e") + "}"
        if name == "d":
            pr = el.find(M + "dPr")
            beg = DELIM.get(_val(pr, "begChr", "("), "(")
            end = DELIM.get(_val(pr, "endChr", ")"), ")")
            sep = _val(pr, "sepChr", "|")
            parts = [self.conv(c) for c in el if c.tag == M + "e"]
            sep = r"\middle" + (DELIM.get(sep, sep) or "|")
            return r"\left%s %s \right%s" % (beg, sep.join(parts), end)
        if name == "func":
            fn = self.kids(el, "fName")
            if re.fullmatch(r"\\mathrm\{(\w+)\}", fn):
                fn = r"\operatorname{%s}" % fn[8:-1]
            return fn + " " + self.kids(el, "e")
        if name == "rad":
            deg = self.kids(el, "deg")
            hide = _val(el.find(M + "radPr"), "degHide") in ("1", "on", "true")
            return (r"\sqrt" if hide or not deg else r"\sqrt[%s]" % deg) + self.arg(el, "e")
        if name == "acc":
            return ACCENT.get(_val(el.find(M + "accPr"), "chr", "\u0302"), r"\hat") + self.arg(el, "e")
        if name == "bar":
            pos = _val(el.find(M + "barPr"), "pos", "bot")
            return (r"\overline" if pos == "top" else r"\underline") + self.arg(el, "e")
        if name == "groupChr":
            pr = el.find(M + "groupChrPr")
            top = _val(pr, "pos", "bot") == "top"
            return (r"\overbrace" if top else r"\underbrace") + self.arg(el, "e")
        if name == "limLow":
            return r"\mathop{%s}\limits_%s" % (self.kids(el, "e"), self.arg(el, "lim"))
        if name == "limUpp":
            return r"\mathop{%s}\limits^%s" % (self.kids(el, "e"), self.arg(el, "lim"))
        if name == "m":
            rows = [" & ".join(self.conv(c) for c in r if c.tag == M + "e")
                    for r in el if r.tag == M + "mr"]
            return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(rows)
        if name == "eqArr":
            rows = [self.conv(c) for c in el if c.tag == M + "e"]
            return r"\begin{gathered}%s\end{gathered}" % r" \\ ".join(rows)
        return self.kids(el)


def omml_to_tex(omath, tracked="accept"):
    return Converter(tracked).conv(omath).strip()


# ----------------------------------------------------------------------------
# Typesetting
# ----------------------------------------------------------------------------

def find_engine():
    """Path of a TeX engine that can load system fonts, or None."""
    extra = ["/Library/TeX/texbin", "/opt/homebrew/bin", "/usr/local/bin"]
    path = os.pathsep.join([os.environ.get("PATH", "")] + extra)
    for name in ("tectonic", "xelatex", "lualatex"):
        p = shutil.which(name, path=path)
        if p:
            return p
    return None


PREAMBLE = r"""\documentclass[%(size)spt]{article}
\usepackage{amsmath}
\usepackage[active,tightpage]{preview}
\setlength\PreviewBorder{0pt}
\usepackage{unicode-math}
\setmainfont{%(stem)s}[Path=%(dir)s/, Extension=.otf, UprightFont=*-Regular,
  ItalicFont=*-It, BoldFont=*-Bold, BoldItalicFont=*-BoldIt]
\setmathfont{%(math)s}
\setmathfont{%(stem)s-Regular.otf}[Path=%(dir)s/, range={up/{latin,Latin,num,greek,Greek},"2B,"3D,"28,"29,"5B,"5D,"2212,"2C,"2E,"3A,"3B,"21,"2F,"27}]
\setmathfont{%(stem)s-It.otf}[Path=%(dir)s/, range={it/{latin,Latin,greek,Greek}}]
\setmathfont{%(stem)s-Bold.otf}[Path=%(dir)s/, range={bfup/{latin,Latin,num,greek,Greek}}]
\setmathfont{%(stem)s-BoldIt.otf}[Path=%(dir)s/, range={bfit/{latin,Latin,greek,Greek}}]
\setmathfont{%(math)s}[range={}]
\newsavebox\eqbox
\newcommand\eq[1]{\sbox\eqbox{$\displaystyle #1$}%%
  \typeout{EQBOX \the\ht\eqbox\space\the\dp\eqbox\space\the\wd\eqbox}%%
  \begin{preview}\usebox\eqbox\end{preview}}
\begin{document}
"""


def _compile(engine, workdir, tex_list, font_dir, stem, size, math_font):
    src = PREAMBLE % {"size": size, "dir": font_dir, "stem": stem, "math": math_font}
    src += "".join("\\eq{%s}\n" % t for t in tex_list) + "\\end{document}\n"
    with open(os.path.join(workdir, "equations.tex"), "w", encoding="utf-8") as fh:
        fh.write(src)
    if os.path.basename(engine) == "tectonic":
        cmd = [engine, "--keep-logs", "--chatter", "minimal", "equations.tex"]
    else:
        cmd = [engine, "-interaction=nonstopmode", "-halt-on-error", "equations.tex"]
    try:
        r = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True, timeout=600)
    except (OSError, subprocess.TimeoutExpired) as e:
        return None, str(e)
    log = ""
    if os.path.exists(os.path.join(workdir, "equations.log")):
        with open(os.path.join(workdir, "equations.log"), encoding="utf-8", errors="replace") as fh:
            log = fh.read()
    boxes = [tuple(float(v) for v in m) for m in
             re.findall(r"EQBOX ([\d.]+)pt ([\d.]+)pt ([\d.]+)pt", log)]
    pdf = os.path.join(workdir, "equations.pdf")
    if r.returncode != 0 or len(boxes) != len(tex_list) or not os.path.exists(pdf):
        err = re.findall(r"^! .*$", log + r.stdout + r.stderr, re.M)
        return None, (err[0] if err else (r.stderr or r.stdout).strip()[-300:])
    return boxes, None


def render(tex_list, out_dir, font_file, size=10, math_font="FiraMath-Regular.otf"):
    """Typeset equations; returns ([{"page","height","depth","width"} or None per
    equation], problem or None). Heights are TeX points above the baseline."""
    if not tex_list:
        return [], None
    engine = find_engine()
    if not engine:
        return [None] * len(tex_list), "no TeX engine (install one with: brew install tectonic)"
    font_dir = os.path.dirname(font_file)
    stem = re.sub(r"-Regular$", "", os.path.splitext(os.path.basename(font_file))[0])
    work = os.path.join(out_dir, "equations")
    os.makedirs(work, exist_ok=True)
    boxes, err = _compile(engine, work, tex_list, font_dir, stem, size, math_font)
    if boxes is None:   # find which equations fail, keep the rest
        ok = []
        for t in tex_list:
            b, _ = _compile(engine, work, [t], font_dir, stem, size, math_font)
            ok.append(b is not None)
        good = [t for t, k in zip(tex_list, ok) if k]
        boxes, err2 = _compile(engine, work, good, font_dir, stem, size, math_font) if good \
            else (None, err)
        if boxes is None:
            return [None] * len(tex_list), err2 or err
        it = iter(boxes)
        boxes = [next(it) if k else None for k in ok]
        err = "could not typeset %d equation(s): %s" % (ok.count(False), err)
    pt = 72.0 / 72.27   # TeX points -> PostScript points
    out, page = [], 0
    pdf = os.path.join(out_dir, "equations.pdf")
    shutil.copyfile(os.path.join(work, "equations.pdf"), pdf)
    for b in boxes:
        if b is None:
            out.append(None)
            continue
        page += 1
        out.append({"file": pdf, "page": page, "height": round(b[0] * pt, 3),
                    "depth": round(b[1] * pt, 3), "width": round(b[2] * pt, 3)})
    return out, err
