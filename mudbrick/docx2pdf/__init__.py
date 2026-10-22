"""Lay out a Word manuscript as a two-column preprint PDF in InDesign.

Writes <manuscript>_preprint.pdf and an editable <manuscript>_preprint.indd next to
the .docx (or where --out says), plus <manuscript>_figure_legends.docx. Figures come
from the figure folder; legends from a captions .docx next to the manuscript or its
FIGURE LEGENDS section. See mudbrick/docx2pdf/README.md for the conventions.
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

    build = tempfile.mkdtemp(prefix="mudbrick_docx2pdf_")
    try:
        manifest = convert.convert(docx, build, cfg, figs)
        print("\n".join(convert.report(manifest)))
        if args.dry_run:
            return 0
        os.makedirs(os.path.dirname(out), exist_ok=True)
        print("laying out in InDesign ...", flush=True)
        print(adobe.run_jsx("indesign", LAYOUT, [build, out]))
        print("PDF:      " + out)
        print("InDesign: " + os.path.splitext(out)[0] + ".indd")
        return 0
    except adobe.AdobeError as e:
        print("mudbrick docx2pdf: InDesign failed: %s" % e)
        return 1
    finally:
        if args.keep_build:
            print("build folder: " + build)
        else:
            shutil.rmtree(build, ignore_errors=True)
