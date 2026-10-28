"""Render a manuscript's Word equations to PDF for docx2pdf (needs pandoc and pdflatex).

    mudbrick docx2pdf paper.docx --render-math       # writes equations/ next to paper.docx

Each display equation, and each inline equation too complex to set as text, is converted
from Word's OMML to LaTeX by pandoc and typeset with pdflatex in Computer Modern at the
size in SIZE_PT. The content is reproduced as written in Word; only Word-specific layout
artefacts (alignment spaces, single-row equation arrays, script/bold letter forms pandoc
cannot express) are cleaned up. Output: one <key>.pdf per equation plus equations.json
with its width, height above the baseline and depth below it, in points.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from . import omml, convert

SIZE_OPTION = "9.5pt"     # Computer Modern at 10.95 pt matches Myriad Pro 10 pt by x-height
BORDER_BP = 1.0          # white margin around each rendered equation (keeps overhangs)
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
BP_PER_PT = 72 / 72.27
MAX_DISPLAY_BP = 208     # widest display equation that sits in a column beside its number
SCALE_NOT_BREAK = 0.75   # a wider display equation that fits when scaled to at least this
                         # is kept on one line (layout.jsx scales it); 1 = always break
GREEK = ("alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa "
         "lambda mu nu xi pi varpi rho varrho sigma varsigma tau upsilon phi varphi chi psi "
         "omega Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega").split()


def equations_in(docx):
    """[(key, element)] for every equation in a .docx, in document order."""
    root = ET.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    return [(omml.key(el), el) for el in root.iter() if el.tag == omml.M + "oMath"]


def needed(docx_paths):
    """Equations that docx2pdf will show as graphics: display ones and complex inline ones."""
    out = {}
    for path in docx_paths:
        cfg = dict(convert.DEFAULT_CONFIG)
        convert.MATH.clear()
        convert.read_paragraphs(path, cfg["tracked_changes"])
        info = dict(convert.MATH)
        for k, el in equations_in(path):
            m = info.get(k)
            if m and (m["display"] or not m["runs"]):
                out[k] = (el, m["display"], m["text"])
    return out


def to_latex(eqs, work):
    """pandoc OMML -> LaTeX for each equation, via one small .docx."""
    ET.register_namespace("w", W_NS)
    ET.register_namespace("m", omml.MNS)
    body = []
    for k, (el, _, _) in eqs.items():
        body.append('<w:p><w:r><w:t xml:space="preserve">EQKEY%s </w:t></w:r>%s</w:p>'
                    % (k, ET.tostring(el, encoding="unicode")))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="%s" xmlns:m="%s"><w:body>%s</w:body></w:document>'
           % (W_NS, omml.MNS, "".join(body)))
    path = os.path.join(work, "eqs.docx")
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", doc)
    ast = json.loads(subprocess.run(["pandoc", path, "-t", "json"], check=True,
                                    capture_output=True, text=True).stdout)
    out = {}
    for block in ast["blocks"]:
        inl = block.get("c") if block.get("t") == "Para" else None
        if not inl or inl[0].get("t") != "Str" or not inl[0]["c"].startswith("EQKEY"):
            continue
        k = inl[0]["c"][5:]
        out[k] = " ".join(x["c"][1] for x in inl if x.get("t") == "Math")
    return out


def _unwrap_array(t):
    """A one-row \\begin{array}{r}...\\end{array} is how Word numbers an equation."""
    m = re.fullmatch(r"\\begin\{array\}\{[lcr]\}\s*(.*?)\s*\\end\{array\}", t, re.S)
    if not m:
        return t
    inner = m.group(1)
    depth, top_level_break = 0, False
    for tok in re.finditer(r"\\begin\{|\\end\{|\\\\", inner):
        if tok.group(0) == "\\begin{":
            depth += 1
        elif tok.group(0) == "\\end{":
            depth -= 1
        elif depth == 0:
            top_level_break = True
    return t if top_level_break else inner


SPACE = r"(?:\s|\\ |\\quad|\\qquad|~)"


def _group(t, i):
    """(content, end) of the balanced {...} group starting at t[i] == "{"."""
    depth = 0
    for j in range(i, len(t)):
        if t[j] == "\\":
            continue
        if t[j] == "{" and (j == 0 or t[j - 1] != "\\"):
            depth += 1
        elif t[j] == "}" and t[j - 1] != "\\":
            depth -= 1
            if depth == 0:
                return t[i + 1:j], j + 1
    raise ValueError("unbalanced braces")


def _braces(t):
    r"""Word's labelled brace (a brace character stacked under a term, the label under
    that) comes out of pandoc as \underset{label}{\overset{term}{<brace>}}; LaTeX writes it
    as \underbrace{term}_{label}. An unlabelled one is \overset{term}{<brace>}."""
    for ch, cmd, script in (("︸", r"\underbrace", "_"), ("︷", r"\overbrace", "^")):
        changed = True
        while ch in t and changed:
            changed = False
            for start in [m for m in range(len(t)) if t.startswith("\\underset{", m)]:
                try:
                    label, e1 = _group(t, start + len("\\underset"))
                    if not t.startswith("{\\overset{", e1):
                        continue
                    term, e2 = _group(t, e1 + len("{\\overset"))
                    mark, e3 = _group(t, e2)
                except (ValueError, IndexError):
                    continue
                if mark.strip() == ch and e3 < len(t) and t[e3] == "}":
                    t = t[:start] + "%s{%s}%s{%s}" % (cmd, term, script, label) + t[e3 + 1:]
                    changed = True
                    break
            if changed:
                continue
            for start in [m for m in range(len(t)) if t.startswith("\\overset{", m)]:
                try:
                    term, e2 = _group(t, start + len("\\overset"))
                    mark, e3 = _group(t, e2)
                except (ValueError, IndexError):
                    continue
                if mark.strip() == ch:
                    t = t[:start] + "%s{%s}" % (cmd, term) + t[e3:]
                    changed = True
                    break
    return t


def clean(tex):
    """Remove Word layout artefacts from pandoc's LaTeX without changing the maths."""
    t = re.sub(r"^%s+|%s+$" % (SPACE, SPACE), "", tex)
    t = _unwrap_array(t.strip())
    t = re.sub(r"(?:\\ ){3,}", " ", t)                 # runs of spaces Word uses for alignment
    t = _braces(t)
    t = re.sub(r"^%s+|%s+$" % (SPACE, SPACE), "", t).rstrip("\\").strip()
    t = re.sub(r"\{(?:\s|\\ )*\}", "{}", t)          # groups left empty by that
    # upright words (Word "plain" style: diag, exp, argmax, minimize, degree) that pandoc
    # leaves as bare italic groups such as {diag}, {= argmax}, {\ exp}
    def upright(m):
        lead = m.group(1)
        if "=" in lead:                       # "{= exp}": keep the normal spacing around =
            lead = r"\;=\;"
        return "{%s\\mathrm{%s}%s}" % (lead, m.group(2), m.group(3))
    t = re.sub(r"(?<![A-Za-z\\_^])\{((?:=|\s|\\ )*)([A-Za-z]{2,})((?:\s|\\ )*)\}", upright, t)
    t = re.sub(r"\\mathcal\{([^{}]*)\}",                # Word script l is ℓ
               lambda m: re.sub(r"(?<![A-Za-z\\])l(?![A-Za-z])", r"\\ell ", m.group(1))
               if not re.search(r"[A-Z]", m.group(1)) else "\\mathcal{%s}" % m.group(1).strip(), t)
    t = re.sub(r"\\mathbf\{\\(%s)\}" % "|".join(GREEK), r"\\boldsymbol{\\\1}", t)
    return t


def _top_level(t):
    r"""Positions in t that are outside braces, \left...\right and environments."""
    depth, lr, env, out, i = 0, 0, 0, [], 0
    while i < len(t):
        if t.startswith("\\left", i) and not t[i + 5:i + 6].isalpha():
            lr += 1; i += 5; continue
        if t.startswith("\\right", i) and not t[i + 6:i + 7].isalpha():
            lr -= 1; i += 6; continue
        if t.startswith("\\begin{", i) or t.startswith("\\end{", i):
            env += 1 if t.startswith("\\begin{", i) else -1
            i = t.index("}", i) + 1             # skip the environment name and its brace
            continue
        c = t[i]
        if c == "\\":                      # a command: record where it starts, skip its name
            if depth == 0 and lr == 0 and env == 0:
                out.append(i)
            m = re.match(r"\\[A-Za-z]+|\\.", t[i:])
            i += len(m.group(0)) if m else 1
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        elif depth == 0 and lr == 0 and env == 0:
            out.append(i)
        i += 1
    return set(out)


def _rstrip(t):
    """Strip trailing spaces, including LaTeX control spaces "\\ " (a lone backslash left
    before an inserted "\\\\" would turn it into a literal "\\&")."""
    return re.sub(r"(?:\\ |\s)+$", "", t)


def line_breaks(t):
    """Two-line versions of a display equation that is too wide for a column, best first.
    Only the line breaks change, never the content."""
    out = []
    top = _top_level(t)
    # Word's "A, B" pair of equations separated by a wide space -> one per line
    q = [m for m in re.finditer(r"\\quad\s*\\quad|\\qquad", t) if m.start() in top]
    if q:
        parts, last = [], 0
        for m in q:
            parts.append(_rstrip(t[last:m.start()]).lstrip()); last = m.end()
        parts.append(t[last:].strip())
        if len(parts) > 2:      # several definitions: all on separate lines, then in pairs
            out.append(r"\begin{gathered}%s\end{gathered}" % r"\\".join(parts))
        out.append(r"\begin{gathered}%s\\%s\end{gathered}"
                   % (_rstrip(t[:q[0].start()]), t[q[0].end():].strip()))
    # long sum: break before a top-level + or - , continuation indented under the =
    eq = next((i for i in sorted(top) if t[i] == "=" ), None)
    ops = [i for i in sorted(top) if t[i] in "+-" and eq is not None and i > eq + 2]
    ops.sort(key=lambda i: abs((i - eq) - (len(t) - eq) / 2))
    for i in ops[:3]:
        out.append(r"\begin{aligned}%s&=%s\\&\quad%s\end{aligned}"
                   % (_rstrip(t[:eq]), _rstrip(t[eq + 1:i]), t[i:]))
    # no "=": break before a top-level + anyway
    if eq is None:
        for i in [i for i in sorted(top) if t[i] == "+"][:2]:
            out.append(r"\begin{aligned}&%s\\&\quad%s\end{aligned}" % (_rstrip(t[:i]), t[i:]))
    return out


DOC = r"""\documentclass[%s,border=%.2fbp]{standalone}
\usepackage{amsmath,amssymb,bm}
\begin{document}
\sbox0{$%s%s$}\typeout{EQDIM \the\wd0;\the\ht0;\the\dp0;}\usebox0
\end{document}
"""


def _compile(work, k, tex, display):
    """Typeset one equation; returns its metrics (bp) and PDF path, or an error message."""
    src = os.path.join(work, k + ".tex")
    with open(src, "w", encoding="utf-8") as fh:
        fh.write(DOC % (SIZE_OPTION, BORDER_BP, "\\displaystyle " if display else "", tex))
    r = subprocess.run(["pdflatex", "-interaction=nonstopmode", "-halt-on-error", k + ".tex"],
                       cwd=work, capture_output=True, text=True)
    try:
        log = open(os.path.join(work, k + ".log"), encoding="latin-1").read()
    except OSError:
        log = ""
    dim = re.search(r"EQDIM (-?[\d.]+)pt;(-?[\d.]+)pt;(-?[\d.]+)pt;", log)
    if r.returncode or not dim:
        err = re.search(r"^! (.*)$", log, re.M)
        return "LaTeX error (%s)" % (err.group(1) if err else "?")
    wd, ht, dp = (float(x) * BP_PER_PT for x in dim.groups())
    pdf = os.path.join(work, k + ".ok.pdf")
    shutil.copyfile(os.path.join(work, k + ".pdf"), pdf)
    return {"w": round(wd + 2 * BORDER_BP, 3), "h": round(ht + BORDER_BP, 3),
            "d": round(dp + BORDER_BP, 3), "pdf": pdf}


def render(docx, out_dir=None, extra_docx=()):
    """Render every equation docx2pdf needs; returns (count, problems)."""
    for tool in ("pandoc", "pdflatex"):
        if not shutil.which(tool):
            raise RuntimeError("%s is not installed (needed to render equations)" % tool)
    out_dir = out_dir or os.path.join(os.path.dirname(os.path.abspath(docx)), "equations")
    os.makedirs(out_dir, exist_ok=True)
    eqs = needed([docx] + list(extra_docx))
    work = tempfile.mkdtemp(prefix="mudbrick_math_")
    problems, index = [], {}
    try:
        latex = to_latex(eqs, work)
        for k, (el, display, text) in eqs.items():
            tex = clean(latex.get(k, ""))
            if not tex:
                problems.append("%s: pandoc could not convert %r" % (k, text[:40]))
                continue
            best, best_pdf = None, os.path.join(work, k + ".best.pdf")
            for attempt, variant in enumerate([tex] + (line_breaks(tex) if display else [])):
                res = _compile(work, k, variant, display)
                if isinstance(res, str):
                    if attempt == 0:
                        problems.append("%s: %s in %s" % (k, res, tex[:60]))
                        break
                    continue
                if best is None or res["w"] < best["w"]:
                    shutil.copyfile(res["pdf"], best_pdf)   # each try overwrites res["pdf"]
                    best = dict(res, latex=variant, pdf=best_pdf)
                if not display or res["w"] <= MAX_DISPLAY_BP:
                    break
                if attempt == 0 and MAX_DISPLAY_BP / res["w"] >= SCALE_NOT_BREAK:
                    break                       # slightly too wide: scale rather than break
            if best is None:
                continue
            shutil.copyfile(best.pop("pdf"), os.path.join(out_dir, k + ".pdf"))
            index[k] = dict(best, file=k + ".pdf", display=display, word_text=text,
                            word_latex=tex, line_broken=best["latex"] != tex)
        with open(os.path.join(out_dir, "equations.json"), "w", encoding="utf-8") as fh:
            json.dump(index, fh, indent=1, ensure_ascii=False)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return len(index), problems
