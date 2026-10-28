# mudbrick 🧱

**Run Adobe apps from the command line.** Each mudbrick tool turns a routine job you'd do
by hand in an Adobe app into one command.

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

### docx2pdf: set up your folder

Put the manuscript and its figures together like this:

```
my_paper/
├── paper.docx              ← the manuscript
├── paper_figures.docx      ← optional: figure captions (see below)
└── figures/
    ├── figure_1/
    │   └── overview_v3.ai  ← any file name works inside a figure_N folder
    ├── figure_2/
    │   └── maps_v2.ai
    └── figure_3/
        └── results.ai
```

- **One folder per main figure,** named `figure_1`, `figure_2`, … (`Figure 1` or
  `fig1` also work). The folder's number decides which figure it is, so the file names
  inside don't matter.
- **Use Illustrator `.ai` files if you can.** Text and lines stay sharp in print. PDF
  also works, and so do PNG, TIFF and JPG. When saving from Illustrator, keep
  **"Create PDF Compatible File"** checked (it's on by default), or InDesign can't read
  the file.
- **Fit the artboard to the artwork** (Object → Artboards → Fit to Artwork Bounds).
  Anything outside the artboard gets cut off. Extra white space is trimmed for you.
- **Leftovers in a figure folder are fine.** If a folder holds several versions, the
  highest one wins (`_v3` beats `_v2`), and `.ai` or `.pdf` beats an image of the same
  figure. Files with "panel" or "subfig" in the name, Illustrator temp files, and
  subfolders such as `prev/` or `old/` are ignored.
- **No folders?** Files named `figure1.ai`, `Fig2_maps.ai` and so on, sitting next to
  the manuscript, work too.
- **Captions** come from a separate Word file next to the manuscript whose name contains
  "figure", "legend" or "caption", or else from a **FIGURE LEGENDS** section at the end
  of the manuscript. Start each caption with `Figure 1.` (bold title), then the panels
  `(A) … (B) …`. Supplementary figures are left out.
- Word comments are ignored. Tracked changes are read as if accepted.
- **Equations:** if the manuscript has Word equations, add `--render-math` to the first run
  (and after editing an equation). It typesets them with LaTeX into an `equations/` folder
  next to the manuscript; needs `pandoc` and `pdflatex`. Simple inline math is set as text.

### docx2pdf: run it

Do a dry run first. It lists the file chosen for each figure and flags problems
such as a missing caption or a placeholder title, without opening InDesign.

```bash
cd my_paper
mudbrick docx2pdf paper.docx -f figures/ --dry-run  # check: which file it picked for each figure, plus warnings
mudbrick docx2pdf paper.docx -f figures/            # make paper_preprint.pdf (+ .indd) next to paper.docx
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
