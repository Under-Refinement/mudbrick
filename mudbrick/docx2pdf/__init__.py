"""Lay out a Word manuscript as a two-column preprint PDF in InDesign.

Writes <manuscript>_preprint.pdf and an editable <manuscript>_preprint.indd next to
the .docx (or where --out says), plus <manuscript>_figure_legends.docx.

Next to the manuscript it also looks for:
  - a captions .docx (name has figure/legend/caption): its captions are used, and its
    embedded pictures stand in for figures that have no file in the figure folder
  - a supplementary .docx (name has supp/supplement/SI): laid out as a separate
    <manuscript>_supplement.pdf (+ .indd) with its figures and tables
See mudbrick/docx2pdf/README.md for the conventions.
"""
import json
import os
import shutil
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
    p.add_argument("--supplement", metavar="DOCX",
                   help="supplementary .docx (default: found by name next to the manuscript)")
    p.add_argument("--no-supplement", action="store_true",
                   help="don't make the supplementary PDF")
    p.add_argument("--dry-run", action="store_true",
                   help="only read the manuscript and figures and report; don't run InDesign")
    p.add_argument("--keep-build", action="store_true",
                   help="keep the intermediate tagged-text folder and print its path")


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
    supp = None if args.no_supplement else (
        os.path.abspath(args.supplement) if args.supplement else convert.find_supplement_docx(docx))
    base = os.path.splitext(out)[0]
    supp_out = (base[:-len("_preprint")] if base.endswith("_preprint") else base) + "_supplement.pdf"

    build = tempfile.mkdtemp(prefix="mudbrick_docx2pdf_")
    supp_build = os.path.join(build, "supplement")
    try:
        manifest = convert.convert(docx, build, cfg, figs, write_outputs=not args.dry_run)
        print("\n".join(convert.report(manifest)))
        smanifest = None
        if supp:
            smanifest = convert.convert_supplement(supp, docx, supp_build, cfg,
                                                   manifest["running_title"])
            print("\n".join(convert.report_supplement(smanifest, supp)))
        if args.dry_run:
            return 0
        os.makedirs(os.path.dirname(out), exist_ok=True)
        print("laying out in InDesign ...", flush=True)
        print(adobe.run_jsx("indesign", LAYOUT, [build, out]))
        print("PDF:      " + out)
        print("InDesign: " + os.path.splitext(out)[0] + ".indd")
        if smanifest:
            print("laying out the supplement in InDesign ...", flush=True)
            print(adobe.run_jsx("indesign", LAYOUT, [supp_build, supp_out]))
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
