# mudbrick 🧱

**Run Adobe apps from the command line.** Each mudbrick tool turns a routine job you'd do
by hand in an Adobe app into one command. (An *adobe* is a sun-dried mud brick.)

```bash
mudbrick docx2pdf paper.docx
```
That command turns a Word manuscript and its figures into a print-ready, two-column
preprint PDF, laid out by InDesign.

---

## Get started (3 steps)

**1. Check you have:**
- macOS, with the Adobe app the tool uses (InDesign for `docx2pdf`)
- Python 3.9 or newer (check with `python3 --version`)

**2. Download mudbrick and put the command on your PATH:**
```bash
git clone https://github.com/Under-Refinement/mudbrick.git
cd mudbrick
ln -s "$PWD/bin/mudbrick" /opt/homebrew/bin/mudbrick   # or any folder on your PATH
```

**3. Check it works:**
```bash
mudbrick --help
```

The first time a tool opens an Adobe app, macOS asks whether Terminal may control
it. Click **OK**.

---

## Tools

| Command | What it does | Needs |
|---|---|---|
| `mudbrick docx2pdf paper.docx` | Turns a Word manuscript and its figures into a preprint PDF. It also saves an editable `.indd` file and a figure-legends `.docx`. [Details →](mudbrick/docx2pdf/README.md) | InDesign, and the Myriad Pro fonts (free at [fonts.adobe.com](https://fonts.adobe.com/fonts/myriad)) |

You can also write a tool as a flag, e.g. `mudbrick --docx2pdf paper.docx`. Each tool
has its own help: `mudbrick docx2pdf --help`.

### Example: docx2pdf

```bash
mudbrick docx2pdf paper.docx                        # figures are taken from paper.docx's folder
mudbrick docx2pdf paper.docx -f figures/ -o out.pdf # choose the figure folder and output file
mudbrick docx2pdf paper.docx --dry-run              # check figures and warnings without opening InDesign
```

---

## For developers

- **Run the tests:** `python3 -m unittest discover tests`. The tests need neither
  InDesign nor real papers.
- **Add a tool:**
  1. Create `mudbrick/<tool>/__init__.py` with `HELP`, `add_arguments(parser)` and
     `run(args)`.
  2. Register it in `TOOLS` in `mudbrick/cli.py`.
  3. Run any Adobe script through `mudbrick.adobe.run_jsx(app, script, args)`.
- mudbrick uses only the Python standard library, so there's nothing to `pip install`.
- Never commit manuscripts, figures or outputs. `.gitignore` blocks `*.docx`, `*.pdf`,
  `*.indd`, `*.ai` and `*.tif`.
