#!/usr/bin/env python3
"""Convert a manuscript .docx into InDesign Tagged Text + a JSON manifest.

The InDesign script layout.jsx reads the output folder and builds the preprint PDF.
Only the Python standard library is used. Normally run through `mudbrick docx2pdf`;
standalone:  python3 convert.py manuscript.docx out_dir [figure_dir] [config.json]

Expected manuscript conventions (as in a typical Word submission):
  * first paragraphs: title, author line, affiliations, notes, "Keywords:",
    "Running title:"; section headings are bold ALL-CAPS paragraphs
  * sub-headings are paragraphs that are entirely bold
  * a "FIGURE LEGENDS" section holds legends "Figure 1. ..." followed by
    panel paragraphs "(A) ..."; supplementary legends (Figure S1, Movie S1)
    are dropped
"""
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

DEFAULT_CONFIG = {
    # Sections (ALL-CAPS bold headings in the docx) that are dropped from the PDF.
    "skip_sections": ["SUPPLEMENTARY TABLES", "SUPPLEMENTARY FIGURES",
                      "SUPPLEMENTARY MATERIALS", "SUPPLEMENTARY INFORMATION"],
    # Short sections moved after the references, in this order.
    "back_matter_order": ["DATA AVAILABILITY", "AUTHOR CONTRIBUTIONS",
                          "COMPETING INTERESTS", "ACKNOWLEDGEMENTS",
                          "ACKNOWLEDGMENTS", "FUNDING"],
    # Front-matter lines dropped (matched at start of paragraph, case-insensitive).
    "drop_front_matter": ["Classifications"],
    # Font file used to detect characters the body font cannot display.
    "body_font_file": "/Applications/Adobe InDesign 2026/Resources/Required/fonts/MyriadPro-Regular.otf",
    # Break the title onto a new line after its colon.
    "title_break_after_colon": True,
    # Italicise the keyword list after "Keywords:".
    "italic_keywords": True,
    # Set "et al." in italics wherever it appears.
    "italic_et_al": True,
    # Crop white margins off figure files (pixels lighter than this are "white").
    "trim_figures": True,
    "trim_threshold": 245,
    "trim_padding": 0.01,  # keep this fraction of the image around the content
    # Word comments are always ignored. Tracked changes: the markup is ignored and the
    # document is read as currently edited ("accept"); "reject" reads the text as it
    # was before the tracked edits.
    "tracked_changes": "accept",
    # Separate Word file with the figure captions (e.g. a figures+captions document).
    # Default: any other .docx next to the manuscript whose name mentions
    # figure/legend/caption. Its captions replace a FIGURE LEGENDS section.
    "legends_docx": None,
}

# Accepted figure formats, most preferred first (vector beats raster for print).
FIG_EXT = (".pdf", ".ai", ".eps", ".psd", ".tif", ".tiff", ".png", ".jpg", ".jpeg")

ENVELOPE = "✉"
LINE_BREAK = " "
# characters InDesign handles itself (spaces, breaks) — never flagged as missing
SPACES = set("         ​"
             "   \t\n‑")
FONT_CHARS = None  # code points available in the body font (None = unknown)
ITALIC_ET_AL = True
WARNINGS = []  # things the author should look at, printed at the end

# Headings that start the body; everything before the first of them is front matter.
BODY_START = {"ABSTRACT", "SUMMARY", "INTRODUCTION", "BACKGROUND", "RESULTS",
              "MAIN TEXT", "SIGNIFICANCE", "SIGNIFICANCE STATEMENT"}
# Front-matter labels, inline ("Keywords: a, b") or as a heading with content below.
LABEL = re.compile(r"^(key ?words|running title|short title)\s*:?\s*(.*)$", re.I | re.S)


# ----------------------------------------------------------------------------
# docx reading
# ----------------------------------------------------------------------------

def _on(el, tag):
    """True/False if a toggle property (w:b, w:i, ...) is set on rPr el, else None."""
    if el is None:
        return None
    x = el.find(W + tag)
    if x is None:
        return None
    return x.get(W + "val", "true") not in ("0", "false", "none")


def load_styles(z):
    """Map styleId -> run toggles (b, i) defined by paragraph/character styles."""
    styles = {}
    try:
        root = ET.fromstring(z.read("word/styles.xml"))
    except KeyError:
        return styles
    for st in root.findall(W + "style"):
        rpr = st.find(W + "rPr")
        based = st.find(W + "basedOn")
        styles[st.get(W + "styleId")] = {
            "b": _on(rpr, "b"), "i": _on(rpr, "i"),
            "based": based.get(W + "val") if based is not None else None,
        }
    return styles


def style_prop(styles, sid, prop):
    seen = set()
    while sid and sid in styles and sid not in seen:
        seen.add(sid)
        v = styles[sid][prop]
        if v is not None:
            return v
        sid = styles[sid]["based"]
    return None


def read_paragraphs(path, tracked="accept"):
    """Paragraphs of a .docx as {"runs": [(text, fmt)]}. Comments are never read;
    tracked changes are applied ("accept") or undone ("reject") without markup."""
    z = zipfile.ZipFile(path)
    styles = load_styles(z)
    body = ET.fromstring(z.read("word/document.xml")).find(W + "body")
    paras = []
    for p in body:  # top-level paragraphs only (tables are skipped)
        if p.tag != W + "p":
            continue
        ppr = p.find(W + "pPr")
        pstyle = None
        if ppr is not None:
            ps = ppr.find(W + "pStyle")
            pstyle = ps.get(W + "val") if ps is not None else None
        runs = []
        in_instr = False  # inside a field's instruction (e.g. Zotero JSON)
        added = {id(r) for tag in ("ins", "moveTo") for c in p.iter(W + tag)
                 for r in c.iter(W + "r")}
        for r in p.iter(W + "r"):
            if tracked == "reject" and id(r) in added:
                continue
            fc = r.find(W + "fldChar")
            if fc is not None:
                t = fc.get(W + "fldCharType")
                if t == "begin":
                    in_instr = True
                elif t in ("separate", "end"):
                    in_instr = False
                continue
            if in_instr:
                continue
            rpr = r.find(W + "rPr")
            if rpr is not None and _on(rpr, "vanish"):
                continue
            rs = rpr.find(W + "rStyle") if rpr is not None else None
            rs = rs.get(W + "val") if rs is not None else None

            def prop(tag):
                v = _on(rpr, tag)
                if v is None and rs:
                    v = style_prop(styles, rs, tag)
                if v is None and pstyle:
                    v = style_prop(styles, pstyle, tag)
                return bool(v)

            va = rpr.find(W + "vertAlign") if rpr is not None else None
            u = rpr.find(W + "u") if rpr is not None else None
            fonts = rpr.find(W + "rFonts") if rpr is not None else None
            font = fonts.get(W + "ascii") if fonts is not None else None
            fmt = {
                "b": prop("b"), "i": prop("i"),
                "sup": va is not None and va.get(W + "val") == "superscript",
                "sub": va is not None and va.get(W + "val") == "subscript",
                "u": u is not None and u.get(W + "val", "single") != "none",
                "mono": font in ("Courier New", "Courier", "Consolas", "Menlo"),
            }
            text = ""
            for ch in r:
                if ch.tag == W + "t" or (ch.tag == W + "delText" and tracked == "reject"):
                    text += ch.text or ""
                elif ch.tag == W + "tab":
                    text += "\t"
                elif ch.tag == W + "br" and ch.get(W + "type") not in ("page", "column"):
                    text += "\n"
                elif ch.tag == W + "noBreakHyphen":
                    text += "‑"
                elif ch.tag == W + "sym":
                    text += chr(int(ch.get(W + "char"), 16))
            if text:
                runs.append((text, fmt))
        paras.append({"runs": merge_runs(runs)})
    return paras


def merge_runs(runs):
    out = []
    for t, f in runs:
        if out and out[-1][1] == f:
            out[-1] = (out[-1][0] + t, f)
        else:
            out.append((t, f))
    return out


def ptext(p):
    return "".join(t for t, _ in p["runs"]).strip()


def all_bold(p):
    # ignore runs that are only punctuation/space (e.g. an unbolded final period)
    rs = [(t, f) for t, f in p["runs"] if re.search(r"\w", t)]
    return bool(rs) and all(f["b"] for _, f in rs)


def is_h1(p):
    t = ptext(p)
    return (all_bold(p) and 0 < len(t) < 60 and t.upper() == t
            and re.search("[A-Z]", t) is not None)


def is_h2(p):
    t = ptext(p)
    return all_bold(p) and 0 < len(t) < 250 and not is_h1(p)


# ----------------------------------------------------------------------------
# Characters
# ----------------------------------------------------------------------------

def read_cmap(path):
    """Code points covered by an OpenType/TrueType font (cmap formats 4 and 12)."""
    data = open(path, "rb").read()
    num = struct.unpack(">H", data[4:6])[0]
    off = None
    for i in range(num):
        tag, _, o, _ = struct.unpack(">4sIII", data[12 + 16 * i:28 + 16 * i])
        if tag == b"cmap":
            off = o
    cps = set()
    for i in range(struct.unpack(">H", data[off + 2:off + 4])[0]):
        _, _, so = struct.unpack(">HHI", data[off + 4 + 8 * i:off + 12 + 8 * i])
        t = off + so
        fmt = struct.unpack(">H", data[t:t + 2])[0]
        if fmt == 4:
            seg = struct.unpack(">H", data[t + 6:t + 8])[0] // 2
            ends = struct.unpack(">%dH" % seg, data[t + 14:t + 14 + 2 * seg])
            starts = struct.unpack(">%dH" % seg, data[t + 16 + 2 * seg:t + 16 + 4 * seg])
            for a, b in zip(starts, ends):
                if a != 0xFFFF:
                    cps.update(range(a, b + 1))
        elif fmt == 12:
            for g in range(struct.unpack(">I", data[t + 12:t + 16])[0]):
                a, b, _ = struct.unpack(">III", data[t + 16 + 12 * g:t + 28 + 12 * g])
                cps.update(range(a, b + 1))
    return cps


def fix_typography(text):
    """Replace look-alike characters that the body font lacks."""
    # modifier-letter primes used as single quotes: opening after space/bracket
    text = re.sub("(?:(?<=^)|(?<=[\\s(\\[{—–-]))[ʹʼ]", "‘", text)
    text = text.replace("ʹ", "’").replace("ʼ", "’")
    # 5′ / 3′ ends -> 5’ / 3’ (Myriad has no prime glyph)
    return text.replace("′", "’").replace("″", "”")


def missing(c):
    return (FONT_CHARS is not None and ord(c) >= 128 and c not in SPACES
            and c != ENVELOPE and ord(c) not in FONT_CHARS)


def esc(s):
    return s.replace("\\", "\\\\").replace("<", "\\<").replace(">", "\\>")


def utf16_units(c):
    b = c.encode("utf-16-be")
    return [int.from_bytes(b[i:i + 2], "big") for i in range(0, len(b), 2)]


# ----------------------------------------------------------------------------
# Tagged-text writer
# ----------------------------------------------------------------------------

class Tagged:
    def __init__(self):
        self.paras = []
        self.cstyles = set()
        self.missing = set()

    def para(self, pstyle, runs, bold_flag="b"):
        runs = [(t, f) for t, f in runs if t]
        if runs:
            runs[0] = (runs[0][0].lstrip(), runs[0][1])
            runs[-1] = (runs[-1][0].rstrip(), runs[-1][1])
        if ITALIC_ET_AL:
            runs = italicize(runs, r"\bet al\.")
        out = "<pstyle:%s>" % pstyle
        for text, f in runs:
            text = fix_typography(text).replace("\n", LINE_BREAK)
            text = re.sub(ENVELOPE + r"(?=\w)", ENVELOPE + " ", text)
            # group characters by whether they need a special font
            for piece in re.findall("%s|[^%s]+" % (ENVELOPE, ENVELOPE), text):
                if piece == ENVELOPE:
                    out += self._run(piece, f, bold_flag, False, env=True)
                    continue
                buf, buf_missing = "", False
                for c in piece:
                    m = missing(c)
                    if m:
                        self.missing.add(c)
                    if buf and m != buf_missing:
                        out += self._run(buf, f, bold_flag, buf_missing)
                        buf = ""
                    buf += c
                    buf_missing = m
                if buf:
                    out += self._run(buf, f, bold_flag, buf_missing)
        self.paras.append(out)

    def _run(self, text, f, bold_flag, fallback, env=False):
        flags = []
        if f.get("b"):
            flags.append(bold_flag)
        for k in ("i", "sup", "sub", "u", "mono"):
            if f.get(k):
                flags.append(k)
        if fallback:
            flags.append("sym")
        if env:
            flags.append("env")
        return self._tag(flags, text)

    def _tag(self, flags, text):
        if not flags:
            return esc(text)
        name = "cs_" + "_".join(flags)
        self.cstyles.add(name)
        return "<cstyle:%s>%s<cstyle:>" % (name, esc(text))

    def write(self, path):
        # ASCII tagged text; every non-ASCII character becomes <0xXXXX>.
        # (InDesign 2026 does not recognise the <UNICODE-UTF8> header.)
        text = "\r".join(self.paras)
        enc = "".join(c if ord(c) < 128 else "<0x000A>" if c == LINE_BREAK else "".join(
            "<0x%04X>" % u for u in utf16_units(c)) for c in text)
        with open(path, "w", encoding="ascii", newline="") as fh:
            fh.write("<ASCII-MAC>\r<vsn:8>\r" + enc)


# ----------------------------------------------------------------------------
# Figure files
# ----------------------------------------------------------------------------

FIG_SKIP = re.compile(r"(^[~.]|\.tmp$|panel|subfig|sub_fig)", re.I)
OLD_DIRS = re.compile(r"^(prev|old|archive|previous|backup|unused)", re.I)


def _usable(name):
    return name.lower().endswith(FIG_EXT) and not FIG_SKIP.search(name)


def _version(path):
    v = re.findall(r"v(\d+)", os.path.splitext(os.path.basename(path))[0], re.I)
    return max(int(x) for x in v) if v else 0


def best_figure_file(paths):
    """Highest version number, then vector over raster, then most recently modified."""
    return min(paths, key=lambda p: (-_version(p),
                                     FIG_EXT.index(os.path.splitext(p)[1].lower()),
                                     -os.path.getmtime(p)))


def find_figure(fig_dir, n):
    """File for main figure n, searched shallowest folder first (down to two levels).
    At each level a folder named figure_N / Fig N wins over file names; otherwise files
    named fig1..., Figure_1..., Fig1_... are used. Old/prev/archive folders are skipped."""
    num = r"[ _\-]*0*%d(?![0-9])" % n
    folder_pat = re.compile(r"^fig(ure)?%s$" % num, re.I)
    file_pat = re.compile(r"^fig(ure)?%s" % num, re.I)
    level = [fig_dir]
    for _ in range(3):
        named, subdirs = [], []
        for root in level:
            entries = sorted(os.listdir(root))
            dirs = [d for d in entries if os.path.isdir(os.path.join(root, d))
                    and not d.startswith(".") and not OLD_DIRS.match(d)]
            for d in dirs:
                if folder_pat.match(d):
                    folder = os.path.join(root, d)
                    inside = [os.path.join(folder, x) for x in os.listdir(folder) if _usable(x)]
                    if inside:
                        return best_figure_file(inside)
            named += [os.path.join(root, x) for x in entries
                      if file_pat.match(x) and _usable(x) and os.path.isfile(os.path.join(root, x))]
            subdirs += [os.path.join(root, d) for d in dirs]
        if named:
            return best_figure_file(named)
        level = subdirs
    return None


def content_box(path, out_dir, threshold, pad):
    """Fractional [left, top, right, bottom] of the non-white area of an image,
    measured on a small BMP rendered by macOS `sips` (None if unavailable)."""
    stem = os.path.join(out_dir, "trim_" + re.sub(r"\W", "_", os.path.basename(path)))
    bmp, src = stem + ".bmp", path
    if path.lower().endswith(".ai"):
        # Illustrator files are PDFs inside, but sips only renders them under a .pdf name
        src = stem + ".pdf"
        shutil.copyfile(path, src)
    try:
        subprocess.run(["sips", "-s", "format", "bmp", "--resampleWidth", "500",
                        src, "--out", bmp], check=True, capture_output=True)
        d = open(bmp, "rb").read()
    except (OSError, subprocess.CalledProcessError):
        return None
    finally:
        if src != path and os.path.exists(src):
            os.remove(src)
    off = struct.unpack("<I", d[10:14])[0]
    w, h = struct.unpack("<ii", d[18:26])
    bpp = struct.unpack("<H", d[28:30])[0] // 8
    if bpp < 3:
        return None
    top_down, h = h < 0, abs(h)
    stride = (w * bpp + 3) // 4 * 4
    x0, y0, x1, y1 = w, h, -1, -1
    for r in range(h):
        row = d[off + r * stride: off + r * stride + w * bpp]
        y = r if top_down else h - 1 - r
        for x in range(w):
            px = row[x * bpp: x * bpp + 3]
            if bpp == 4 and row[x * bpp + 3] == 0:  # transparent
                continue
            if min(px) < threshold:
                x0, x1 = min(x0, x), max(x1, x)
                y0, y1 = min(y0, y), max(y1, y)
    os.remove(bmp)
    if x1 < 0:
        return None
    return [max(0, x0 / w - pad), max(0, y0 / h - pad),
            min(1, (x1 + 1) / w + pad), min(1, (y1 + 1) / h + pad)]


def cited_figures(paras):
    """Main-figure numbers cited in the text (not "Supplementary Figure N" / "Figure S1")."""
    nums = set()
    text = " ".join(ptext(p) for p in paras)
    for m in re.finditer(r"Fig(?:ure)?s?\.?\s+((?:S?\d+[A-Za-z]*(?:\s*[-\u2013,]\s*|\s+and\s+)?)+)", text):
        if re.search(r"(Supplementa(l|ry)|Extended Data|Appendix)\s*$", text[max(0, m.start() - 20):m.start()]):
            continue
        nums |= {int(x) for x in re.findall(r"(?<![S\d])(\d+)", m.group(1))}
    return nums


def add_figures_without_legends(manifest, fig_dir, out_dir, cited):
    """Cited figures that have a file but no legend still get placed, with a visible note."""
    have = {f["num"] for f in manifest["figures"]}
    for n in sorted(cited):
        if n in have or not fig_dir or not find_figure(fig_dir, n):
            continue
        tg = Tagged()
        tg.para("LegendTitle", [("Figure %d." % n, {"b": True})], bold_flag="sb")
        tg.para("Legend", [("[Legend not found in the manuscript.]", {"i": True})])
        finish_legend({"num": n, "tg": tg}, out_dir, manifest)
        WARNINGS.append("Figure %d has no legend in the manuscript; placeholder used." % n)
    manifest["figures"].sort(key=lambda f: f["num"])


def attach_figures(manifest, fig_dir, out_dir, cfg):
    for f in manifest["figures"]:
        path = find_figure(fig_dir, f["num"]) if fig_dir else None
        f["file"] = path
        f["crop"] = [0, 0, 1, 1]
        if not path:
            WARNINGS.append("No file found for Figure %d." % f["num"])
        elif cfg["trim_figures"]:
            f["crop"] = content_box(path, out_dir, cfg["trim_threshold"],
                                    cfg["trim_padding"]) or f["crop"]


# ----------------------------------------------------------------------------
# Manuscript structure
# ----------------------------------------------------------------------------

def italicize(runs, pattern):
    """Split runs so that matches of pattern are italic."""
    out = []
    for t, f in runs:
        if f.get("i"):
            out.append((t, f))
            continue
        pos = 0
        for m in re.finditer(pattern, t):
            out.append((t[pos:m.start()], f))
            out.append((m.group(0), dict(f, i=True)))
            pos = m.end()
        out.append((t[pos:], f))
    return merge_runs([(t, f) for t, f in out if t])


def break_after_colon(runs):
    """Replace the first ': ' in the runs by ':' + forced line break."""
    out, done = [], False
    for t, f in runs:
        if not done and ":" in t:
            head, _, tail = t.partition(":")
            t = head + ":\n" + tail.lstrip()
            done = True
            out.append((t, f))
            continue
        if done and out and out[-1][0].endswith("\n"):
            t = t.lstrip()
        out.append((t, f))
    return out


def plain(runs):
    return [(t, dict(f, b=False)) for t, f in runs]


def convert(docx_path, out_dir, cfg, fig_dir=None):
    global FONT_CHARS, ITALIC_ET_AL
    del WARNINGS[:]
    ITALIC_ET_AL = cfg["italic_et_al"]
    if cfg.get("body_font_file") and os.path.exists(cfg["body_font_file"]):
        FONT_CHARS = read_cmap(cfg["body_font_file"])
    paras = read_paragraphs(docx_path, cfg["tracked_changes"])
    os.makedirs(out_dir, exist_ok=True)
    manifest = {"running_title": "", "figures": [], "cstyles": []}

    # ---- split into front matter and headed sections --------------------
    def head_name(p):
        return ptext(p).rstrip(":").strip().upper()
    first_text = next(i for i, p in enumerate(paras) if ptext(p))
    first_h1 = next((i for i, p in enumerate(paras) if is_h1(p) and head_name(p) in BODY_START),
                    None)
    if first_h1 is None:
        first_h1 = next(i for i, p in enumerate(paras)
                        if i > first_text and is_h1(p) and not LABEL.match(ptext(p)))
    front, gap = [], False
    for p in paras[:first_h1]:
        if ptext(p):
            p["gap"] = gap
            front.append(p)
            gap = False
        else:
            gap = True
    sections = []  # [heading, [paras incl. blanks]]
    for p in paras[first_h1:]:
        if is_h1(p):
            sections.append([ptext(p).rstrip(":").strip(), []])
        else:
            sections[-1][1].append(p)

    # ---- front matter -----------------------------------------------------
    title = Tagged()
    title_runs = list(front[0]["runs"])
    if re.fullmatch(r"\[.*\]", ptext(front[0])):
        WARNINGS.append("Title is a placeholder: %s" % ptext(front[0]))
    if cfg["title_break_after_colon"]:
        title_runs = break_after_colon(title_runs)
    title.para("Title", title_runs)
    title.para("Authors", front[1]["runs"])
    affil_seen = False
    label, keywords, running = None, [], []
    for p in front[2:]:
        t = ptext(p)
        low = t.lower()
        if any(low.startswith(d.lower()) for d in cfg["drop_front_matter"]):
            label = "drop"
            continue
        m = LABEL.match(t)
        if m:
            label = "keywords" if m.group(1).lower().startswith("key") else "running"
            if m.group(2).strip():
                (keywords if label == "keywords" else running).append(m.group(2).strip())
            continue
        if label:  # content paragraph under a label heading
            if label != "drop":
                (keywords if label == "keywords" else running).append(t)
            continue
        runs = list(p["runs"])
        # "1Department of Biology," -> "1. Department of Biology"
        if runs and runs[0][1]["sup"] and runs[0][0].strip().isdigit():
            runs[0] = (runs[0][0].strip() + ". ", dict(runs[0][1], sup=False))
            runs[-1] = (runs[-1][0].rstrip().rstrip(","), runs[-1][1])
        if t.startswith("*") or t.startswith(ENVELOPE):
            style = "NoteGap" if p["gap"] else "Note"
            if runs[0][0].strip() == ENVELOPE:  # full-size icon followed by a space
                runs[0] = (ENVELOPE + " ", dict(runs[0][1], sup=False))
                runs[1:] = [(runs[1][0].lstrip(), runs[1][1])] + runs[2:] if len(runs) > 1 else []
        elif not affil_seen:
            style = "AffilFirst"
        else:
            style = "AffilGap" if p["gap"] else "Affil"
        affil_seen = True
        title.para(style, runs)
    if running:
        manifest["running_title"] = running[0]
        if len(running) > 1:
            WARNINGS.append("%d running titles given; using the first: %s" % (len(running), running[0]))
    else:
        WARNINGS.append("No running title found; footer left without one.")

    # ---- abstract, body, references, back matter --------------------------
    abstract, body = Tagged(), Tagged()
    refs_sec, back = None, {}
    for head, ps in sections:
        H = head.upper()
        if H == "ABSTRACT":
            abstract.para("AbstractHead", [(head, {})])
            for p in ps:
                if ptext(p):
                    abstract.para("Abstract", p["runs"])
            if keywords:
                abstract.para("Keywords", [("Keywords: ", {}), (
                    ", ".join(k.rstrip(".") for k in keywords) + ".",
                    {"i": cfg["italic_keywords"]})])
        elif H in cfg["skip_sections"]:
            continue
        elif H in ("FIGURE LEGENDS", "FIGURES"):
            parse_legends(ps, out_dir, manifest)
        elif H in ("REFERENCES", "LITERATURE CITED", "BIBLIOGRAPHY"):
            refs_sec = (head, ps)
        elif H in cfg["back_matter_order"]:
            back[H] = (head, ps)
        else:
            emit_section(body, head, ps)
    if refs_sec:
        head, ps = refs_sec
        body.para("H1", [(head, {})])
        for p in ps:
            if ptext(p):
                body.para("Reference", p["runs"])
    for name in cfg["back_matter_order"]:
        if name in back:
            emit_section(body, *back[name], first_style="H1Back")

    title.write(os.path.join(out_dir, "title.txt"))
    abstract.write(os.path.join(out_dir, "abstract.txt"))
    body.write(os.path.join(out_dir, "body.txt"))

    legends_doc = find_legends_docx(docx_path, cfg)
    if legends_doc:
        if manifest["figures"]:
            WARNINGS.append("Captions taken from %s instead of the manuscript's legends section."
                            % os.path.basename(legends_doc))
        manifest["figures"] = []
        parse_legends(read_paragraphs(legends_doc, cfg["tracked_changes"]), out_dir,
                      manifest, strict=True)
        manifest["legends_source"] = legends_doc
    if any(f.get("_legend") for f in manifest["figures"]):
        out = os.path.splitext(os.path.abspath(docx_path))[0] + "_figure_legends.docx"
        write_legends_docx(out, sorted(manifest["figures"], key=lambda f: f["num"]))
        manifest["legends_docx_written"] = out
    add_figures_without_legends(manifest, fig_dir, out_dir, cited_figures(paras))
    cs, miss = set(), set()
    for f in manifest["figures"]:
        f.pop("_legend", None)
    for tg in [title, abstract, body] + [f.pop("_tg") for f in manifest["figures"]]:
        cs |= tg.cstyles
        miss |= tg.missing
    manifest["cstyles"] = sorted(cs)
    attach_figures(manifest, fig_dir, out_dir, cfg)
    manifest["missing_glyphs"] = sorted("U+%04X" % ord(c) for c in miss)
    manifest["warnings"] = WARNINGS
    with open(os.path.join(out_dir, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=1, ensure_ascii=False)
    return manifest


def emit_section(tg, head, ps, first_style="H1"):
    """Body paragraphs get a blank line after them; runs of short one-liners
    (e.g. author-contribution lines) are kept tight."""
    tg.para(first_style, [(head, {})])
    content = [p for p in ps if ptext(p)]
    for n, p in enumerate(content):
        if is_h2(p):
            tg.para("H2", p["runs"])
            continue
        if re.fullmatch(r"\[[^\]]*\]", ptext(p)):
            WARNINGS.append("Editorial note left in the text: %s" % ptext(p))
        nxt = content[n + 1] if n + 1 < len(content) else None
        short = len(ptext(p)) < 90 and nxt is not None and len(ptext(nxt)) < 90 \
            and not is_h2(nxt)
        tg.para("BodyTight" if short else "Body", p["runs"])


FIG_START = re.compile(r"(Figure|Fig\.?)\s+(S?)(\d+)\s*[.:]?")
PANEL = re.compile(r"\((?:[A-Z](?:\s*[,\-\u2013]\s*[A-Z])*)\)")


def legend_blocks(paras, strict=False):
    """[(n, [runs of each paragraph])] for main figures. A legend starts with
    "Figure N." (or "Figure N:") and continues with the following paragraphs; with
    strict=True (captions document) only paragraphs starting with a panel label
    such as "(B)" are taken, so notes between figures are left out."""
    blocks, cur = [], None
    for p in paras:
        t = ptext(p)
        if not t:
            continue
        m = FIG_START.match(t)
        if m or re.match(r"(Movie|Video|Table)\s+S?\d+", t):
            cur = None
            if m and not m.group(2) and "continued" not in t.lower():
                cur = (int(m.group(3)), [p["runs"]])
                blocks.append(cur)
            continue
        if cur and (not strict or PANEL.match(t)):
            cur[1].append(p["runs"])
        elif strict:
            cur = None
    return blocks


def normalize_legend(paragraphs):
    """[(style, runs)]: the bold title as its own LegendTitle paragraph, then one
    Legend paragraph per panel. A bold panel label like "(B)" or "(D,E)" starts a new
    panel only at the start of a sentence (not "...picks (B) and..."). "Figure 4:"
    becomes "Figure 4."."""
    out = []
    for k, runs in enumerate(paragraphs):
        pieces = []
        for t, f in runs:  # split bold text in front of each panel label
            parts = re.split(r"(?=%s)" % PANEL.pattern, t) if f.get("b") else [t]
            pieces += [(x, f) for x in parts if x]
        cur, in_title = [], k == 0
        for t, f in pieces:
            so_far = "".join(x for x, _ in cur).strip()
            at_start = not so_far or so_far[-1] in ".!?"
            label = f.get("b") and PANEL.match(t.lstrip())
            if in_title and (not f.get("b") and re.search(r"\w", t) or label and so_far):
                out.append(("LegendTitle", cur))
                cur, in_title = [], False
            elif label and at_start and so_far:
                out.append(("Legend", cur))
                cur = []
            cur.append((t, f))
        if cur:
            out.append(("LegendTitle" if in_title else "Legend", cur))
    if out and out[0][0] == "LegendTitle":
        t0, f0 = out[0][1][0]
        out[0][1][0] = (re.sub(r"^(\s*(?:Figure|Fig\.?)\s+\d+)\s*:", r"\1.", t0), f0)
    return [(st, merge_runs(r)) for st, r in out]


def parse_legends(ps, out_dir, manifest, strict=False):
    for n, paragraphs in legend_blocks(ps, strict):
        if any(f["num"] == n for f in manifest["figures"]):
            continue
        tg = Tagged()
        legend = normalize_legend(paragraphs)
        for style, runs in legend:
            tg.para(style, runs, bold_flag="sb")
        finish_legend({"num": n, "tg": tg, "legend": legend}, out_dir, manifest)


def finish_legend(cur, out_dir, manifest):
    fn = "legend_%d.txt" % cur["num"]
    cur["tg"].write(os.path.join(out_dir, fn))
    manifest["figures"].append({"num": cur["num"], "legend": fn, "_tg": cur["tg"],
                                "_legend": cur.get("legend")})


def find_legends_docx(docx_path, cfg):
    if cfg.get("legends_docx"):
        return cfg["legends_docx"]
    folder, me = os.path.dirname(os.path.abspath(docx_path)), os.path.basename(docx_path)
    found = [x for x in os.listdir(folder)
             if x.lower().endswith(".docx") and x != me and not x.startswith("~$")
             and not x.endswith("_figure_legends.docx")
             and re.search(r"figure|legend|caption", x, re.I)]
    if len(found) > 1:
        WARNINGS.append("Several caption documents found; using the newest of %s" % found)
    return max((os.path.join(folder, x) for x in found), key=os.path.getmtime) if found else None


def write_legends_docx(path, figures):
    """A plain Word file with a FIGURE LEGENDS section in the house format: bold
    "Figure N. Title." paragraph, then one paragraph per panel with a bold label."""
    def run(t, f):
        props = "".join(tag for k, tag in (("b", "<w:b/>"), ("i", "<w:i/>"),
                        ("sup", '<w:vertAlign w:val="superscript"/>'),
                        ("sub", '<w:vertAlign w:val="subscript"/>')) if f.get(k))
        t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (props, t)

    def para(runs):
        return "<w:p>%s</w:p>" % "".join(run(t, f) for t, f in runs)
    body = [para([("FIGURE LEGENDS", {"b": True})])]
    for fig in figures:
        for style, runs in fig["_legend"] or []:
            if style == "LegendTitle":
                runs = [(t, dict(f, b=True)) for t, f in runs]
            body.append(para(runs))
        body.append("<w:p/>")
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml",
                   '<?xml version="1.0" encoding="UTF-8"?><w:document %s><w:body>%s</w:body></w:document>'
                   % (ns, "".join(body)))


def report(m):
    """Lines summarising a conversion: figures found, legends written, warnings."""
    lines = ["running title: " + m["running_title"]]
    for f in m["figures"]:
        lines.append("figure %d: %s" % (
            f["num"], os.path.basename(f["file"]) if f["file"] else "NOT FOUND"))
    if m.get("legends_docx_written"):
        lines.append("figure legends written to: " + m["legends_docx_written"])
    if m["missing_glyphs"]:
        lines.append("characters not in body font (set in fallback font): %s"
                     % ", ".join(m["missing_glyphs"]))
    lines += ["WARNING: " + w for w in m["warnings"]]
    return lines


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    cfg = dict(DEFAULT_CONFIG)
    fig_dir = None
    for extra in sys.argv[3:]:
        if extra.lower().endswith(".json"):
            with open(extra) as fh:
                cfg.update(json.load(fh))
        else:
            fig_dir = extra
    print("\n".join(report(convert(sys.argv[1], sys.argv[2], cfg, fig_dir))))
