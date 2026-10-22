# docx2pdf: Word manuscript → InDesign → print-ready PDF

`mudbrick docx2pdf` turns a Word manuscript and its figure files into a two-column
preprint PDF in the Davis Lab bioRxiv style: Myriad Pro type, red section heads, a
running-title footer, and full-width figures with legends. It uses your local copy of
InDesign and also saves an editable `.indd` file.

## Usage

```bash
mudbrick docx2pdf manuscript.docx [-f figure_folder] [-o output.pdf] [-c settings.json]
mudbrick docx2pdf manuscript.docx --dry-run    # check parsing, figures and warnings only
```

* `-f` defaults to the folder that contains the .docx file.
* `-o` defaults to `manuscript_preprint.pdf`, saved next to the .docx file. An `.indd`
  file with the same name is saved alongside it, and so is
  `manuscript_figure_legends.docx`.
* InDesign has to be installed. The first time you run it, macOS asks you to let
  Terminal control InDesign; allow it.

A run takes from about 20 seconds to several minutes; very large Illustrator figures
take the longest. The tool prints which file it used for each figure, where it placed
it, and any warnings about the manuscript.

The PDF is exported for print with InDesign's "Preprint - Maximum Quality" preset,
which the script creates:
* images are compressed losslessly (ZIP, never JPEG)
* colours are not converted
* fonts are embedded in full
* it is a PDF 1.6 file, so transparency stays live

Images finer than 600 ppi at their printed size are reduced to 450 ppi. That is more
than any printer resolves, and it keeps files manageable: figures that embed
10,000-px images can otherwise produce PDFs of several hundred MB.
InDesign cannot resample some soft-masked images inside Illustrator files. When that
happens, the script retries more gently (images above 900 ppi reduced to 600 ppi, and
as a last resort no downsampling) and reports which setting it used.

## What the Word file needs

* The title is the first paragraph and the author line is the second. They are
  followed by the affiliations and notes (`*Equal contributions`, `✉…`).
* `Keywords:` and `Running title:` can be written inline (`Keywords: a, b`) or as a
  heading with the content in the paragraphs below. If there are several
  running-title candidates, the first one is used. Empty keywords are left out.
* Section headings are **bold ALL-CAPS** paragraphs, and the body starts at the
  first of these headings: ABSTRACT, INTRODUCTION, RESULTS, … Subheadings are
  paragraphs that are entirely bold.
* Word comments are always ignored. Tracked changes are read as currently edited:
  the markup is ignored and the edits are taken. To read the text as it was before
  the edits instead, set `"tracked_changes": "reject"`.

### Figure legends

Legends are taken from the first of these two places that exists:
1. A **separate captions document**: any other .docx next to the manuscript whose
   name mentions figure/legend/caption (e.g. `…_maintext_figures.docx`). Each caption
   starts with `Figure N.` or `Figure N:`. A note between figures is left out unless
   it starts with a panel label.
2. A `FIGURE LEGENDS` section in the manuscript.

In both cases the legend is normalized:
* the bold title becomes its own paragraph
* each panel starts a new paragraph at a bold label (`(A)`, `(B,C)`) that begins a
  sentence; in-sentence references like "…picks **(B)** and…" stay inline
* `Figure N:` becomes `Figure N.`

The result is also saved as `<manuscript>_figure_legends.docx`, with a FIGURE LEGENDS
section in the house format. A figure that is cited in the text and has a file but no
legend gets a visible "[Legend not found…]" placeholder.

Supplementary legends (`Figure S1`, `Movie S1`), `SUPPLEMENTARY TABLES` and
"Supplementary Figure N" citations are ignored.

### Figure files

These are searched from the figure folder down two levels, nearest first:
* A folder named `figure_N`, `Figure 3` or `fig2` holds figure N, whatever the
  file names inside it are.
* Otherwise, files named `fig1…`, `Figure_2…` or `Fig3_…` are used.

When several files match, they are skipped or ranked as follows:
* skipped: Illustrator temp files (`~ai-…tmp`), files with "panel" or "subfig" in
  the name, subfolders, and `prev/`, `old/` or `archive/` folders
* chosen: the highest version (`_v9` beats `_v2` beats no version), then vector
  (pdf/ai) over raster, then the newest file

White margins, including empty artboard space, are trimmed automatically.

The script prints warnings for things to fix in the manuscript: a placeholder title,
a missing legend, notes like `[jhd stopped here]`, or several running titles.

## Layout rules

* Each figure goes at the top of the page right after the page where it is first
  cited, so it always follows the text that first mentions it. The next section may
  start before it: e.g. DISCUSSION can begin on the page where Results end, with the
  last Results figure at the top of the following page, so no space is left blank. Set
  `sectionAfterFigures: true` in `layout.jsx` to push the section heading to a new page
  below the figure instead.
* The abstract block on page 1 ends at 730 pt (`abstractBottom`), leaving clear space
  above the footer rule.
* Otherwise, each figure goes at the top of the page after the page where it is first cited.
  A figure can shrink to 90% of its file size so that more of its legend fits on the
  same page. Legend panels that still don't fit continue at the foot of the next page,
  under "Figure N (continued from prior page)."
* The references come after the main text, followed by Data Availability, Author
  Contributions, Competing Interests and Acknowledgements.
* "et al." is set in italics, and the title breaks after its colon.
* If the body font lacks a character, the script substitutes a matching character
  (for example ʹ becomes ‘ or ’, and ′ becomes ’) or sets that character in a fallback
  font.

## Settings

* Page geometry, fonts, colours and figure behaviour are set in the `CFG` block at the top
  of `layout.jsx`. For example, `figurePosition: "bottom"` puts figures at the foot of
  the page, and `minFigScale` sets how small a figure may get.
* Manuscript parsing (skipped sections, back-matter order, italic "et al.", figure
  trimming) is set in `DEFAULT_CONFIG` in `convert.py`. You can also override
  these for a single paper with a JSON file:
  `mudbrick docx2pdf paper.docx -c myconfig.json`.

## Known differences from a hand-set layout

* Legend titles use Myriad Pro **Semibold**. If it isn't active, the tool falls back to
  Bold and says so; turn on the Myriad Pro Semibold style at fonts.adobe.com to fix it.
* Artwork outside an Illustrator artboard is cut off; fit the artboard to the artwork
  (Object → Artboards → Fit to Artwork Bounds).
* Copy edits made by hand in InDesign (e.g. trimming URLs from references) aren't
  reproduced. Make these edits in Word, or touch up the saved `.indd`.
* Tables in the main text are not handled yet; supplementary tables are skipped.
