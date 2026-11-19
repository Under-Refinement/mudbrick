"""Lay out a Word manuscript as a two-column preprint PDF in InDesign.

Writes <manuscript>_preprint.pdf and an editable <manuscript>_preprint.indd next to
the .docx (or where --out says), plus <manuscript>_figure_legends.docx.

Next to the manuscript it also looks for:
  - a captions .docx (name has figure/legend/caption): its captions are used, and its
    embedded pictures stand in for figures that have no file in the figure folder
  - a supplementary .docx (name has supp/supplement/SI): laid out as a separate
    <manuscript>_supplement.pdf (+ .indd) with its figures and tables. Its figures come
    from supp_figure_N/ or Supp_FigN files in the figure folder, else its pasted pictures
See mudbrick/docx2pdf/README.md for the conventions.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile

from .. import adobe
from . import convert

HELP = "Word manuscript + figures -> print-ready preprint PDF (InDesign)"
LAYOUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "layout.jsx")


def add_arguments(p):
    p.add_argument("docx", help="manuscript .docx")
    p.add_argument("-f", "--figures", metavar="DIR",
                   help="figure folder (default: the manuscript's folder)")
    p.add_argument("-o", "--out", metavar="PDF",
                   help="output PDF (default: <manuscript>_preprint.pdf next to the .docx)")
    p.add_argument("-c", "--config", metavar="JSON",
                   help="JSON file overriding conversion settings (see README)")
    p.add_argument("--figures-from-captions", action="store_true",
                   help="use the pictures in the captions .docx even if figure files exist")
    p.add_argument("--figure-border", action="store_true",
                   help="draw a thin border around every figure")
    p.add_argument("--link-figures", action="store_true",
                   help="link the figure files from the .indd instead of embedding them")
    p.add_argument("--supplement", metavar="DOCX",
                   help="supplementary .docx (default: found by name next to the manuscript)")
    p.add_argument("--no-supplement", action="store_true",
                   help="don't make the supplementary PDF")
    p.add_argument("--supplement-only", action="store_true",
                   help="only make the supplementary PDF, not the main preprint")
    p.add_argument("--dry-run", action="store_true",
                   help="only read the manuscript and figures and report; don't run InDesign")
    p.add_argument("--keep-build", action="store_true",
                   help="keep the intermediate tagged-text folder and print its path")


def pdf_page_count(path):
    """Pages in an existing PDF (via Spotlight metadata), or None."""
    try:
        out = subprocess.run(["mdls", "-raw", "-name", "kMDItemNumberOfPages", path],
                             capture_output=True, text=True).stdout.strip()
        return int(out)
    except (OSError, ValueError):
        return None


def run(args):
    docx = os.path.abspath(args.docx)
    if not os.path.isfile(docx):
        print("mudbrick docx2pdf: no such file: %s" % args.docx)
        return 2
    figs = os.path.abspath(args.figures) if args.figures else os.path.dirname(docx)
    out = os.path.abspath(args.out) if args.out else os.path.splitext(docx)[0] + "_preprint.pdf"
    cfg = dict(convert.DEFAULT_CONFIG)
    if args.config:
        with open(args.config) as fh:
            cfg.update(json.load(fh))
    if args.figures_from_captions:
        cfg["figures_from_captions"] = True
    if args.figure_border:
        cfg["figure_border"] = True
    if args.link_figures:
        cfg["embed_figures"] = False
    supp = None if args.no_supplement else (
        os.path.abspath(args.supplement) if args.supplement else convert.find_supplement_docx(docx))
    if args.supplement_only and args.no_supplement:
        print("mudbrick docx2pdf: --supplement-only and --no-supplement can't be combined")
        return 2
    if args.supplement_only and not supp:
        print("mudbrick docx2pdf: no supplement found; name it with supp/supplement/SI "
              "or use --supplement FILE")
        return 2
    base = os.path.splitext(out)[0]
    supp_out = (base[:-len("_preprint")] if base.endswith("_preprint") else base) + "_supplement.pdf"

    build = tempfile.mkdtemp(prefix="mudbrick_docx2pdf_")
    supp_build = os.path.join(build, "supplement")
    try:
        if args.supplement_only:   # read the manuscript for its title and footer only
            manifest = convert.convert(docx, build, dict(cfg, trim_figures=False), figs,
                                       write_outputs=False)
            lines = ["WARNING: " + w for w in manifest["warnings"] if "running title" in w]
        else:
            manifest = convert.convert(docx, build, cfg, figs, write_outputs=not args.dry_run)
            lines = convert.report(manifest)
        print("\n".join(lines))
        smanifest = None
        if supp:
            smanifest = convert.convert_supplement(supp, docx, supp_build, cfg,
                                                   manifest["running_title"], figs)
            print("\n".join(convert.report_supplement(smanifest, supp)))
        if args.dry_run:
            return 0
        os.makedirs(os.path.dirname(out), exist_ok=True)
        main_pages = None
        if not args.supplement_only:
            print("laying out in InDesign ...", flush=True)
            log = adobe.run_jsx("indesign", LAYOUT, [build, out])
            print(log)
            print("PDF:      " + out)
            print("InDesign: " + os.path.splitext(out)[0] + ".indd")
            m = re.search(r"pages: (\d+)", log)
            main_pages = int(m.group(1)) if m else None
        elif os.path.exists(out):
            main_pages = pdf_page_count(out)
        if smanifest:
            start = 0
            if cfg["supplement_page_numbers"] == "continue":
                if main_pages:
                    start = main_pages + 1
                else:
                    print("note: main PDF page count unknown; supplement pages numbered %s1, %s2, ..."
                          % (cfg["supplement_page_prefix"], cfg["supplement_page_prefix"]))
            print("laying out the supplement in InDesign ...", flush=True)
            print(adobe.run_jsx("indesign", LAYOUT, [supp_build, supp_out, start]))
            print("PDF:      " + supp_out)
            print("InDesign: " + os.path.splitext(supp_out)[0] + ".indd")
        return 0
    except adobe.AdobeError as e:
        print("mudbrick docx2pdf: InDesign failed: %s" % e)
        return 1
    finally:
        if args.keep_build:
            print("build folder: " + build)
        else:
            shutil.rmtree(build, ignore_errors=True)
