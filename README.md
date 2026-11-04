# mudbrick 🧱

I built mudbrick so I can run Adobe apps from the command line instead of doing the same
layout work by hand. So far it has one tool, `docx2pdf`, which turns a Word manuscript
into a clean, print-ready preprint with one command:

```bash
mudbrick docx2pdf paper.docx
```

InDesign does the layout, so you also get an `.indd` file you can open and tweak.

## Setup

You need a Mac, InDesign, Python 3.9 or newer, and the Myriad Pro fonts (Regular, Italic,
Bold and Semibold). The fonts are free with Creative Cloud at
[fonts.adobe.com](https://fonts.adobe.com/fonts/myriad).

```bash
git clone https://github.com/Under-Refinement/mudbrick.git
cd mudbrick
ln -s "$PWD/bin/mudbrick" /opt/homebrew/bin/mudbrick   # or any folder on your PATH
mudbrick --help                                        # check that it works
```

There's nothing else to install. The first time you run it, your Mac asks if Terminal
can control InDesign. Say yes.

## Making a preprint

**1. Put your files in one folder.** The simplest setup is just Word files:

```
my_paper/
├── paper.docx                  ← the manuscript
├── paper_maintext_figures.docx ← each figure pasted in, with its caption right below it
└── paper_supp_materials.docx   ← optional: supplementary figures, captions and tables
```

File names matter, because that's how mudbrick tells them apart. The captions file needs
"figure", "legend" or "caption" in its name, and the supplement needs "supp",
"supplement" or "SI".

**2. Do a dry run.** It shows which figure files it picked and warns you about problems,
without opening InDesign or saving anything:

```bash
cd my_paper
mudbrick docx2pdf paper.docx --dry-run
```

**3. Make the PDFs:**

```bash
mudbrick docx2pdf paper.docx
```

You'll get these, next to your manuscript:

| File | What it is |
|---|---|
| `paper_preprint.pdf` and `.indd` | The preprint |
| `paper_supplement.pdf` and `.indd` | The supplement, with pages numbered S1, S2, … (only if you have a supplement file) |
| `paper_figure_legends.docx` | All the main figure captions in one Word file |

## Want sharp figures? Use your `.ai` files

Figures pasted into Word are usually only about 200 ppi. That's fine on screen but soft
in print, and mudbrick will warn you. For print, give it your original figures instead,
with one folder per figure:

```
my_paper/
├── paper.docx
├── paper_maintext_figures.docx   ← still used for the captions
└── figures/
    ├── figure_1/
    │   └── overview_v3.ai        ← any file name works in here
    ├── figure_2/
    │   └── maps_v2.ai
    └── figure_3/
        └── results.ai
```

```bash
mudbrick docx2pdf paper.docx -f figures/
```

A few things worth knowing:

- **The folder number is the figure number.** File names inside don't matter.
- **It's fine to leave old versions in there.** The highest version wins (`_v3` beats
  `_v2`), and `.ai` or `.pdf` beats an image. It ignores files with "panel" or "subfig"
  in the name, Illustrator temp files, and folders like `prev/` or `old/`.
- **In Illustrator, fit the artboard to your artwork** (Object → Artboards → Fit to
  Artwork Bounds). Anything outside the artboard gets cut off. Extra white space gets
  trimmed for you.
- **Keep "Create PDF Compatible File" checked** when you save (it's on by default), or
  InDesign can't read the file.
- **Prefer no folders?** Files named `figure1.ai`, `Fig2_maps.ai` and so on, sitting next
  to the manuscript, work too.

## How to write the manuscript

- Start with the title, then the authors, then affiliations and correspondence.
- Add `Keywords: …` and `Running title: …` after that. The running title goes in the
  footer. Without one, the footer uses your title, cut off with "…" if it's too long.
- Section headings are **bold ALL CAPS** (ABSTRACT, INTRODUCTION, …).
- Subheadings are paragraphs that are entirely bold.
- Start each caption with `Figure 1.` and a bold title, then the panels: `(A) … (B) …`.
  Captions can also go in a FIGURE LEGENDS section of the manuscript instead of a
  separate file.
- Word comments are ignored, and tracked changes are read as if you accepted them.

## What it does for you

- **Figures** go at the top of the page right after they're first mentioned, never before.
- **Captions** are justified, with the panels running together in one paragraph. A long
  caption continues on the next page if it has to.
- **Spacing** between paragraphs and above headings is the same everywhere. A heading
  always stays with its text.
- **Equations** that sit on their own line are centered.
- **The title page** tightens up to fit a long title or abstract, and keywords always stay
  on page 1.
- **Data availability, author contributions** and similar sections move after the
  references, whether you wrote them as headings or as bold labels at the start of a
  paragraph.
- **In the supplement,** every figure stays on the same page as its caption and grows to
  the full caption width when it fits. Tables become real InDesign tables in a clean,
  minimal style, and a table that fits on one page never gets split.
- **The PDF** is made for printing: lossless images, fully embedded fonts, and anything
  above 600 ppi reduced to 450 ppi so the files stay a reasonable size.

## Options

| Flag | What it does |
|---|---|
| `-f DIR` | Use the figure files in this folder |
| `-o out.pdf` | Choose where the PDF goes |
| `--dry-run` | Check everything without making PDFs |
| `--supplement FILE` | Pick the supplement file yourself |
| `--no-supplement` | Skip the supplement |
| `--figures-from-captions` | Use the figures pasted into Word even if figure files exist |
| `--figure-border` | Add a thin border around every figure |
| `-c settings.json` | Change other settings for one paper |

You can also write the tool as a flag: `mudbrick --docx2pdf paper.docx`.

The full list of rules and settings is in the
[docx2pdf README](mudbrick/docx2pdf/README.md).

## Adding a new tool

Each tool is its own folder in `mudbrick/`:

1. Create `mudbrick/<tool>/__init__.py` with `HELP`, `add_arguments(parser)` and `run(args)`.
2. Register it in `TOOLS` in `mudbrick/cli.py`, and add it to `packages` in `pyproject.toml`.
3. Run any Adobe scripts through `mudbrick.adobe.run_jsx(app, script, args)`.

To run the tests (they don't need InDesign): `python3 -m unittest discover tests`

Please don't commit manuscripts, figures or PDFs. `.gitignore` blocks them, so they
shouldn't get in by accident.
