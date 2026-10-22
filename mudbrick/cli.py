"""The `mudbrick` command: one subcommand per tool.

    mudbrick docx2pdf paper.docx            # same as: mudbrick --docx2pdf paper.docx
    mudbrick docx2pdf --help

A tool is a module with HELP (one line), add_arguments(parser) and run(args);
register it in TOOLS below.
"""
import argparse
import importlib
import sys

from . import __version__

TOOLS = {
    "docx2pdf": "mudbrick.docx2pdf",
}


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    # allow the flag spelling too: `mudbrick --docx2pdf paper.docx`
    if argv and argv[0].startswith("--") and argv[0][2:] in TOOLS:
        argv[0] = argv[0][2:]

    parser = argparse.ArgumentParser(
        prog="mudbrick",
        description="Command-line tools that drive Adobe apps.")
    parser.add_argument("--version", action="version", version="mudbrick " + __version__)
    sub = parser.add_subparsers(dest="tool", metavar="<tool>")
    modules = {}
    for name, path in TOOLS.items():
        mod = modules[name] = importlib.import_module(path)
        p = sub.add_parser(name, help=mod.HELP, description=mod.__doc__,
                           formatter_class=argparse.RawDescriptionHelpFormatter)
        mod.add_arguments(p)

    args = parser.parse_args(argv)
    if not args.tool:
        parser.print_help()
        return 1
    return modules[args.tool].run(args) or 0


if __name__ == "__main__":
    sys.exit(main())
