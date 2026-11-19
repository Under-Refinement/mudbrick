# mudbrick 🧱

mudbrick was built to run Adobe apps from the command line, so that routine layout work
doesn't have to be done by hand. It currently has one tool, `docx2pdf`, which turns a Word
manuscript into a clean, print-ready preprint with one command:

```bash
mudbrick docx2pdf paper.docx
```

The layout is done by InDesign, so an editable `.indd` file is saved alongside the PDF.
All figures are embedded in it, so it can be shared and opened on any computer
without the original figure files.

## Setup

Required: a Mac, InDesign, Python 3.9 or newer, and the Myriad Pro fonts (Regular, Italic,
Bold and Semibold). The fonts are free with Creative Cloud at
[fonts.adobe.com](https://fonts.adobe.com/fonts/myriad).

```bash
git clone https://github.com/Under-Refinement/mudbrick.git
cd mudbrick
ln -s "$PWD/bin/mudbrick" /opt/homebrew/bin/mudbrick   # or any folder on your PATH
mudbrick --help                                        # check that it works
```

For beautifully typeset equations, a TeX engine is also needed:
`brew install tectonic`. Nothing else needs to be installed. On the first run, macOS asks whether Terminal may
control InDesign; this must be allowed.

## Making a preprint

**1. Put the files in one folder.** The simplest setup is just Word files:

```
my_paper/
├── paper.docx                  ← the manuscript
├── paper_maintext_figures.docx ← each figure pasted in, with its caption right below it
└── paper_supp_materials.docx   ← optional: supplementary figures, captions and tables
```

File names matter, because they are how the files are told apart. The captions file needs
"figure", "legend" or "caption" in its name, and the supplement needs "supp",
"supplement" or "SI".

**2. Do a dry run.** The figure files that were picked are listed and problems are flagged,
without opening InDesign or saving anything:

```bash
cd my_paper
mudbrick docx2pdf paper.docx --dry-run
```

**3. Make the PDFs:**

```bash
mudbrick docx2pdf paper.docx
```

These are saved next to the manuscript:

| File | What it is |
|---|---|
| `paper_preprint.pdf` and `.indd` | The preprint |
| `paper_supplement.pdf` and `.indd` | The supplement, with pages numbered SI-1, SI-2, … (only if there is a supplement file) |
| `paper_figure_legends.docx` | All the main figure captions in one Word file |

## Sharp figures: use `.ai` files

Figures pasted into Word are usually only about 200 ppi, which is fine on screen but soft in
print, and a warning is shown. For print, the original figures should be supplied instead,
with one folder per figure:

```
my_paper/
├── paper.docx
├── paper_maintext_figures.docx   ← still used for the captions
├── paper_supp_materials.docx     ← still used for supplement captions and tables
└── figures/
    ├── figure_1/
    │   └── overview_v3.ai        ← any file name works in here
    ├── figure_2/
    │   └── maps_v2.ai
    ├── supp_figure_1/
    │   └── controls.ai
    └── supp_figure_2/
        └── statistics_v2.ai
```

```bash
mudbrick docx2pdf paper.docx -f figures/
```

- **The folder number is the figure number.** File names inside don't matter.
- **Old versions can stay in the folder.** The highest version wins (`_v3` beats `_v2`), and
  `.ai` or `.pdf` beats an image. Files with "panel" or "subfig" in the name, Illustrator
  temp files, and folders like `prev/` or `old/` are ignored.
- **In Illustrator, the artboard should fit the artwork** (Object → Artboards → Fit to
  Artwork Bounds). Anything outside the artboard gets cut off. Extra white space is trimmed
  automatically.
- **"Create PDF Compatible File" must stay checked** when saving (it's on by default), or
  InDesign can't read the file.
- **Folders are optional.** Files named `figure1.ai` or `Fig2_maps.ai` (main) and
  `Supp_Fig1.ai` or `FigS2.ai` (supplement), sitting next to the manuscript, also work.

## Running on a supplement

No separate command is needed. A Word file with "supp", "supplement" or "SI" in its name,
next to the manuscript, is picked up automatically and becomes `paper_supplement.pdf`:

```bash
mudbrick docx2pdf paper.docx                                # supplement found by name
mudbrick docx2pdf paper.docx --supplement other_file.docx   # or chosen directly
mudbrick docx2pdf paper.docx --supplement-only              # only the supplement, which is faster
mudbrick docx2pdf paper.docx --no-supplement                # or skipped
```

The supplement file should be set up like this:

- **Movies:** a caption starting `Supplementary Movie 1.` (or `Movie S1.`) for each.
- **Figures:** each caption starts with `Supplementary Figure 1.` The figure comes from a
  matching file in the `-f` folder (`supp_figure_1/`, `figure_S1/`, `Supp_Fig1.ai`, or
  `figure_1/` inside a `supplementary/` folder). If there is no file, the picture pasted
  directly above the caption is used. The two can be mixed.
- **Tables:** normal Word tables, each with a caption starting `Supplementary Table 1.`
  (above or below the table in Word), in the supplement file or in a file of their own
  whose name starts with "Table" (e.g. `Tables S1-6.docx`, next to the manuscript or in the
  figure folder; or `--supp-tables FILE`). Column headings should be bold. A bold row that spans the whole
  table (like "Overall performance") becomes a grey section band.

The supplement is laid out like a journal supplement:

- **A cover page:** "Supplementary Information" above the paper's title, authors and
  affiliations, all taken from the manuscript.
- **Sections:** SUPPLEMENTARY MOVIES, FIGURES and TABLES, each starting on a new page
  (only those that exist).
- **Tables** sit in the two-column grid with the caption below. Narrow tables take one
  column and wide ones span both.
- **Figures** come one per page, with the caption below; figure and caption are separate
  frames, so each can be moved on its own in InDesign.
- **Page numbers** are SI-1, SI-2, … To continue from the main preprint instead (if it
  is 14 pages, the supplement starts at 15), set `"supplement_page_numbers": "continue"`
  in a settings file.
- **Labels** follow how the manuscript cites the supplement: "Figure S1." if the text says
  "Figure S1", or "Supplementary Figure 1." if it says "Supplementary Figure 1".

## How the manuscript should be written

- Start with the title, then the authors, then affiliations and correspondence.
- Add `Keywords: …` and `Running title: …` after that. The running title goes in the footer.
  Without one, the footer uses the full title, on two lines if needed (cut off with "…"
  only if even two lines aren't enough).
- Section headings are **bold ALL CAPS** (ABSTRACT, INTRODUCTION, …).
- Subheadings are paragraphs that are entirely bold.
- Each caption starts with `Figure 1.` and a bold title, followed by the panels:
  `(A) … (B) …`. Captions can also go in a FIGURE LEGENDS section of the manuscript instead
  of a separate file.
- Word comments are ignored, and tracked changes are read as if accepted.

## What gets handled automatically

- **Figures** are placed at the top of the page right after they're first mentioned, never
  before.
- **Captions** are justified, with the panels running together in one paragraph. A long
  caption continues on the next page if needed.
- **Spacing** between paragraphs and above headings is the same everywhere, and a heading
  always stays with its text.
- **Equations** made with Word's equation editor are typeset with TeX in Myriad Pro,
  centred, with the equation number on the right (this needs `brew install tectonic`).
- **The title page** tightens up to fit a long title or abstract, and keywords always stay on
  page 1.
- **Data availability, author contributions** and similar sections are moved after the
  references, whether they were written as headings or as bold labels at the start of a
  paragraph.
- **Figure references** in the text ("Figure 1B, S1C") are set in semibold.
- **Figures are never scaled,** so their font sizes stay as drawn: each is trimmed to its
  artwork, centred, and placed at 100%. One wider or taller than the space overprints the
  margins, with a warning.
- **In the supplement,** every figure stays on the same page as its caption. Tables become real InDesign tables in a clean, minimal
  style, and a table that fits on one page is never split.
- **The PDF** is made for printing: lossless images, fully embedded fonts, and anything above
  600 ppi reduced to 450 ppi to keep file sizes reasonable.

## Options

| Flag | What it does |
|---|---|
| `-f DIR` | Use the figure files in this folder |
| `-o out.pdf` | Choose where the PDF goes |
| `--dry-run` | Check everything without making PDFs |
| `--supplement FILE` | Choose the supplement file directly |
| `--supplement-only` | Only make the supplement |
| `--supp-tables FILE` | Take the supplementary tables from this Word file |
| `--no-supplement` | Skip the supplement |
| `--figures-from-captions` | Use the figures pasted into Word even if figure files exist |
| `--figure-border` | Add a thin border around every figure |
| `--link-figures` | Link the figure files from the .indd instead of embedding them |
| `-c settings.json` | Change other settings for one paper |

The tool can also be written as a flag: `mudbrick --docx2pdf paper.docx`.

The full list of rules and settings is in the
[docx2pdf README](mudbrick/docx2pdf/README.md).

## Adding a new tool

Each tool is its own folder in `mudbrick/`:

1. Create `mudbrick/<tool>/__init__.py` with `HELP`, `add_arguments(parser)` and `run(args)`.
2. Register it in `TOOLS` in `mudbrick/cli.py`, and add it to `packages` in `pyproject.toml`.
3. Run any Adobe scripts through `mudbrick.adobe.run_jsx(app, script, args)`.

The tests don't need InDesign: `python3 -m unittest discover tests`

Manuscripts, figures and PDFs should never be committed. They are blocked by `.gitignore`.
