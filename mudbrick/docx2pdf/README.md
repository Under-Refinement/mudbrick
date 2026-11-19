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
* Every figure (`.ai`, `.pdf` or picture) is embedded in the saved `.indd`, so the file
  can be opened and edited on another computer without the figure files. The `.indd`
  is therefore about as large as all the figure files together. Set `embedFigures:
  false` in `layout.jsx` to keep linked figures instead.
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
* The footer shows the running title, set a little below the footer rule. If the
  manuscript gives none, the footer uses the full title, wrapping onto a second
  (balanced) line when needed. Only a title too long for two lines is cut at a word
  boundary and ended with "…", so it never runs into the page number. A warning
  suggests adding a `Running title:`.
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

### Figures pasted into the captions document

If the figure folder has no file for a figure, the tool uses the picture pasted
directly above that figure's caption in the captions document. With
`--figures-from-captions`, it uses the pasted pictures even when files exist. Tracked
changes apply here too: a picture marked as deleted is skipped. When a picture is
coarser than 300 ppi at full width, the tool warns, because a vector `.ai` or `.pdf`
would print sharper.

### Supplementary information

A `.docx` next to the manuscript whose name contains "supp", "supplement" or "SI"
becomes a second PDF, `<manuscript>_supplement.pdf`, with an `.indd` alongside it. Its
layout follows a journal-style supplement:

* **Cover page:** "Supplementary Information" above the main paper's title, authors,
  affiliations and correspondence.
* **Sections:** SUPPLEMENTARY MOVIES, SUPPLEMENTARY FIGURES and SUPPLEMENTARY TABLES
  (setting `supplement_order`), each starting a new page under a red header. Only
  sections with content appear.
* **Movies:** each legend at full width.
* **Tables:**
  * Real InDesign tables in the two-column grid, with the caption below the table. A
    table that Word sized narrower than a column takes one column; wider tables span both
    (setting `supplement_table_columns`).
  * The tables can be in their own Word file: a .docx whose name starts with "Table"
    (e.g. `Tables S1-6.docx`) next to the manuscript or in the figure folder, or chosen
    with `--supp-tables FILE` (setting `supplement_tables_docx`). Tables in the
    supplement file are then ignored.
  * The style is minimal: rules above and below the table and one under the column
    headings; a group heading spanning several columns (e.g. "tRNA" over "P-site" and
    "A-site") gets a short rule under just those columns. No other lines. Row labels are
    left-aligned and the other columns centred. A bold row spanning the table (e.g.
    "Overall performance") becomes a light grey band.
  * Heading rows are the top rows whose first cell is empty or bold; they repeat when a
    table has to break. Merged cells, Word's column widths, line breaks within a cell
    and bold, italic, superscript and subscript text are kept as typed.
  * A table that fits on a page is never split, and stays with its caption. A note
    paragraph after a table is kept with it.
* **Figures:**
  * One per page (setting `supplement_figure_per_page`), with the caption below on the
    same page.
  * A figure is placed at 100% of its size (never scaled, so its font sizes stay as
    drawn), trimmed to its artwork and centred on the text column. If it is wider than
    the column or taller than the room above its caption, it overprints the margins (or
    the caption) and a warning says by how much.
  * The caption's bold title runs into its text (setting `supplement_title_inline`).
* **Page numbers** are SI-1, SI-2, … (setting `supplement_page_prefix`). With
  `"supplement_page_numbers": "continue"` they continue from the main PDF instead (and
  fall back to the prefixed numbers when the main PDF's page count is unknown).
* **Labels** follow the manuscript's citation style: `Figure S1.` / `Table S1.` /
  `Movie S1.` when the text cites "Figure S1", otherwise `Supplementary Figure 1.` and so
  on. Odd original wording ("Table 1:", "Supplementary Figure. 5.") is normalized.

Use `--supplement FILE` to pick the supplement file yourself, `--no-supplement` to skip
it, or `--supplement-only` to make only the supplement.

A supplementary figure comes from a file in the figure folder if one matches:
`supp_figure_1/`, `figure_S1/`, `Supp_Fig1.ai`, `FigS1.ai`, or `figure_1/` inside a
`supp*`/`SI` folder. The same selection rules and white-margin trimming as main figures
apply. Otherwise the picture pasted above its caption is used. The main-figure search
never looks in supplement folders, and a figure's own folder is only ever used for that
figure.

### Title page

If the title, authors and abstract don't fit on page 1, the white space above and
between them shrinks step by step. Only if that isn't enough does the abstract drop to
10 pt. Keywords always stay on page 1.

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
  A figure is placed at 100% of its size, never scaled, so its font sizes stay as drawn.
  Its white margins are trimmed (the Illustrator artboard need not fit the artwork) and
  the artwork is centred on the text column; if it is wider than the column it overprints
  into the side margins and a warning says by how much. Legend panels that don't fit
  under it continue at the foot of the next page, under "Figure N (continued from prior
  page)."
* The paper's text runs in two separate, linked text frames per page (one per column),
  so either column's height can be adjusted by hand in InDesign. The supplement uses one
  two-column frame per page, because its captions and wide tables span both columns.
* Figures have no border. Use `--figure-border`, or `"figure_border": true` in a
  settings file, to draw a thin 0.5 pt border around every figure in both PDFs.
* Figure files are embedded in the saved .indd files, so they open on any computer.
  Use `--link-figures`, or `"embed_figures": false`, to keep them linked to the figure
  files instead, which keeps the .indd small.
* Stand-alone Word equations (made with Insert → Equation, on their own line,
  optionally followed by a number such as `(1)`) are converted to LaTeX and typeset
  with TeX, in Myriad Pro: letters, digits and Greek come from the body font, and only
  the symbols it lacks (large sums, integrals, fraction rules) from Fira Math. Each
  equation is placed as a vector graphic, centred in the column on the text baseline,
  with its number flush right. This needs a TeX engine: `brew install tectonic` (or an
  installed TeX Live with XeLaTeX). Without one, or with `"typeset_equations": false`,
  equations stay as centred text, as do equations typed as plain text (short lines
  with "=" and almost no ordinary words). Equations inside a sentence stay as text.
* Captions (figure and table titles and legends) are justified, like the body text.
  Manual line breaks inside a caption become spaces, so no justified line gets
  stretched. Bold position labels such as **(left)**, **(right)**, **(top)** and
  **(inset)** count as panel labels, like **(A)**.
* A heading always moves to the next column together with its first lines of text.
  If that leaves more than a line empty at the bottom of a column, the gap is shared
  out between that column's ordinary paragraphs, adding at most 12 pt per break.
  Headings keep exactly their usual spacing, so a column with few paragraph breaks may
  still end slightly short. Settings: `fillColumns` and `columnFillLimit` in
  `layout.jsx`.
* All body paragraphs are spaced evenly, 6 pt apart, and every section heading has
  10 pt above it (`paragraphSpacing` and `headingSpaceAbove` in `layout.jsx`).
  Subheadings get no extra space: just the paragraph spacing after body text, and none
  directly under a section heading (`subheadingSpaceAbove`).
* References are 3 pt apart (`referenceSpacing`), and that spacing is never widened to
  fill out a column.
* Figure references in the text ("Figure 1A", "Figures 2A, 3B-D and S4", "Fig. S2",
  "Supplementary Figure 3") are set in semibold, through a "Figure reference" character
  style that can be restyled in InDesign (`figureRefStyle`; `null` leaves them as typed).
  A reference written in lowercase ("figure 3B") gets a warning. To keep runs of short lines, such as an author-contributions list,
  tight instead, set `"uniform_paragraph_spacing": false`.
* A legend's bold title sits on its own line, and the panels then run on in one
  paragraph: "…end of A. **(B)** …". If a legend doesn't fit under its figure, it
  breaks at a line and continues on the next page. To start each panel on its own
  line instead, set `"legend_panels_inline": false`.
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
