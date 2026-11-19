// layout.jsx — build a two-column preprint PDF in InDesign from the output of
// docx_to_indesign.py.
//
// Run from the command line via make_pdf.sh, or from InDesign's Scripts panel
// (it then asks for the build folder and the output PDF).
//
// arguments: [buildDir, outputPdf]
#target indesign

// ---------------------------------------------------------------------------
// Layout settings (points). Measured from the reference InDesign PDF.
// ---------------------------------------------------------------------------
var CFG = {
    pageW: 612, pageH: 792,
    top: 46.8, bottom: 54, inside: 50.4, outside: 54,
    gutter: 12,
    footerRule: 741, footerTop: 746,   // footer text sits 5 pt below the rule
    footerLines: 2,            // running title may wrap to this many lines before shortening
    font: "Myriad Pro",
    fallbackFont: "Arial Unicode MS", // for characters Myriad Pro lacks
    emojiFont: ["Noto Color Emoji SVG", "Regular"],  // for the correspondence envelope
    kerning: "Metrics",
    red: [192, 0, 0],          // section headings
    envelope: [255, 192, 0],   // correspondence icon
    figGap: 8,                 // figure -> legend
    ruleGap: 6,                // legend -> separator rule
    wrapGap: 12,               // separator rule -> body text
    minLegendLines: 5,         // keep at least this much legend with its figure
    paragraphSpacing: 6,       // space between body paragraphs (pt)
    headingSpaceAbove: 10,     // space above section headings (pt)
    subheadingSpaceAbove: 0,   // extra space above subheadings, on top of the paragraph spacing (pt)
    referenceSpacing: 3,       // space between references (pt); never widened to fill columns
    fillColumns: true,         // fill gaps left when a heading moves to the next column
    columnFillLimit: 12,       // most extra space added between two paragraphs to do so (pt)
    figurePosition: "top",     // "top" or "bottom" of the page after first citation
    abstractBottom: 730,       // foot of the abstract block on page 1 (footer rule at 741)
    sectionAfterFigures: false, // true: push a section heading (e.g. DISCUSSION) to a new
                                // page rather than let it start before a figure cited
                                // at the end of the previous section
    printPPI: 450,             // export: images finer than downsampleAbovePPI are
    downsampleAbovePPI: 600,   //   reduced to printPPI (lossless ZIP either way)
    embedFigures: true         // embed every placed .ai/.pdf/image in the saved .indd, so
                               // it opens on any computer without the figure files
};

var BUILD, OUTPDF, START_PAGE = 0;   // START_PAGE: first page number (supplement)
if (typeof arguments !== "undefined" && arguments.length >= 2) {
    BUILD = Folder(arguments[0]); OUTPDF = File(arguments[1]);
    if (arguments.length >= 3) START_PAGE = parseInt(arguments[2], 10) || 0;
} else {
    BUILD = Folder.selectDialog("Folder produced by docx_to_indesign.py");
    OUTPDF = File.saveDialog("Save PDF as", "*.pdf");
}

function readText(f) { f.encoding = "UTF-8"; f.open("r"); var s = f.read(); f.close(); return s; }
var M = eval("(" + readText(File(BUILD + "/manifest.json")) + ")");

app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
var LOG = [];

// ---------------------------------------------------------------------------
// Document, swatches and styles
// ---------------------------------------------------------------------------
var doc = app.documents.add(false);
with (doc.documentPreferences) {
    pageWidth = CFG.pageW; pageHeight = CFG.pageH; facingPages = true; pagesPerDocument = 1;
}
doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
with (doc.marginPreferences) {
    top = CFG.top; bottom = CFG.bottom; left = CFG.inside; right = CFG.outside;
    columnCount = 2; columnGutter = CFG.gutter;
}

function rgb(name, v) {
    var c = doc.colors.itemByName(name);
    if (!c.isValid) c = doc.colors.add({ name: name, model: ColorModel.PROCESS, space: ColorSpace.RGB, colorValue: v });
    return c;
}
var RED = rgb("Heading Red", CFG.red);
var GOLD = rgb("Envelope", CFG.envelope);
var BLACK = doc.swatches.itemByName("Black");

function hasFont(family, style) {
    var f = app.fonts.itemByName(family + "\t" + style);
    return f.isValid && f.status == FontStatus.INSTALLED;
}
var SEMIBOLD = hasFont(CFG.font, "Semibold") ? "Semibold" : "Bold";
if (SEMIBOLD != "Semibold") LOG.push("Myriad Pro Semibold not installed; legend titles use Bold.");

function pstyle(name, props) {
    var s = doc.paragraphStyles.itemByName(name);
    if (!s.isValid) s = doc.paragraphStyles.add({ name: name });
    var base = {
        appliedFont: CFG.font, fontStyle: "Regular", pointSize: 10, leading: 12,
        hyphenation: false, fillColor: BLACK, spaceBefore: 0, spaceAfter: 0,
        justification: Justification.LEFT_ALIGN, kerningMethod: CFG.kerning,
        alignToBaseline: false
    };
    for (var k in props) base[k] = props[k];
    if (base.keepAllLinesTogether) base.keepLinesTogether = true;  // master switch
    s.properties = base;
    return s;
}
var J = Justification.LEFT_JUSTIFIED, C = Justification.CENTER_ALIGN;
pstyle("Body",       { justification: J, spaceAfter: CFG.paragraphSpacing });
pstyle("BodyTight",  { justification: J });
pstyle("Equation",   { justification: C, spaceAfter: CFG.paragraphSpacing });  // stand-alone equations
// TeX-typeset equations: centred on the column by a centre tab, number on a right tab.
var COL_W = (CFG.pageW - CFG.inside - CFG.outside - CFG.gutter) / 2;
(function () {
    var st = pstyle("EquationTeX", { justification: Justification.LEFT_ALIGN,
        spaceAfter: CFG.paragraphSpacing, spaceBefore: 2, keepLinesTogether: true });
    st.tabStops.everyItem().remove();
    st.tabStops.add({ alignment: TabStopAlignment.CENTER_ALIGN, position: COL_W / 2 });
    st.tabStops.add({ alignment: TabStopAlignment.RIGHT_ALIGN, position: COL_W });
})();
// Headings get the same space above them everywhere (paragraph spacing + this);
// InDesign drops it when a heading starts a column.
var HEAD_GAP = Math.max(0, CFG.headingSpaceAbove - CFG.paragraphSpacing);
pstyle("H1",         { fontStyle: "Bold", pointSize: 12, leading: 14.4, fillColor: RED, keepWithNext: 2, keepAllLinesTogether: true, spaceBefore: HEAD_GAP });
pstyle("H1Back",     { fontStyle: "Bold", pointSize: 12, leading: 14.4, fillColor: RED, keepWithNext: 2, keepAllLinesTogether: true, spaceBefore: HEAD_GAP });
pstyle("H2",         { fontStyle: "Bold", keepWithNext: 2, keepAllLinesTogether: true, spaceBefore: CFG.subheadingSpaceAbove });
pstyle("Reference",  { justification: J, pointSize: 8, leading: 9.6, spaceAfter: CFG.referenceSpacing,
                       keepAllLinesTogether: true });
pstyle("Title",      { fontStyle: "Bold", pointSize: 16, leading: 19.2, justification: C });
pstyle("Authors",    { fontStyle: "Bold", pointSize: 12, leading: 14.4, justification: C, spaceBefore: 54 });
pstyle("AffilFirst", { justification: C, spaceBefore: 56 });
pstyle("Affil",      { justification: C });
pstyle("AffilGap",   { justification: C, spaceBefore: 12 });
pstyle("Note",       { justification: C });
pstyle("NoteGap",    { justification: C, spaceBefore: 14 });
pstyle("AbstractHead", { fontStyle: "Bold", pointSize: 12, leading: 14.4, fillColor: RED, keepWithNext: 2, keepAllLinesTogether: true });
pstyle("Abstract",   { justification: J, pointSize: 11, leading: 13.2 });
pstyle("Keywords",   { pointSize: 11, leading: 13.2, spaceBefore: 26.4 });
pstyle("LegendTitle", { fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8, justification: J });
pstyle("Legend",     { pointSize: 9, leading: 10.8, spaceBefore: 2, justification: J,
                       keepLinesTogether: true, keepFirstLines: 2, keepLastLines: 2 });
pstyle("LegendCont", { fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8, justification: J });
pstyle("Footer",     { fontStyle: "Italic", justification: C, balanceRaggedLines: true });
pstyle("PageNum",    { fontStyle: "Bold", justification: Justification.RIGHT_ALIGN });

// Character styles are named after their flags: cs_b_i_sup -> bold italic superscript.
function cstyle(name) {
    var s = doc.characterStyles.itemByName(name);
    if (!s.isValid) s = doc.characterStyles.add({ name: name });
    var flags = name.replace(/^cs_/, "").split("_"), has = {};
    for (var i = 0; i < flags.length; i++) has[flags[i]] = true;
    var p = {};
    var weight = has.b ? "Bold" : (has.sb ? SEMIBOLD : "");
    if (weight || has.i) p.fontStyle = weight ? (has.i ? weight + " Italic" : weight) : "Italic";
    if (has.sup) p.position = Position.SUPERSCRIPT;
    if (has.sub) p.position = Position.SUBSCRIPT;
    if (has.u) { p.underline = true; p.underlineOffset = 1.5; }
    if (has.mono) { p.appliedFont = "Courier New"; if (!p.fontStyle) p.fontStyle = "Regular"; }
    if (has.sym) { p.appliedFont = CFG.fallbackFont; if (!p.fontStyle || !hasFont(CFG.fallbackFont, p.fontStyle)) p.fontStyle = "Regular"; }
    if (has.env) {  // colour emoji font if available, else a gold dingbat
        if (hasFont(CFG.emojiFont[0], CFG.emojiFont[1])) { p.appliedFont = CFG.emojiFont[0]; p.fontStyle = CFG.emojiFont[1]; }
        else { p.appliedFont = "Zapf Dingbats"; p.fontStyle = "Regular"; p.fillColor = GOLD; }
    }
    s.properties = p;
}
for (var ci = 0; ci < M.cstyles.length; ci++) cstyle(M.cstyles[ci]);

// Supplementary-information mode: a cover page, then a two-column story in sections
// (movies, tables, figures). Headers, figures, captions and wide tables span both
// columns; tables narrower than a column sit in one.
var SI = M.mode == "supplement";
function span(props) {
    var o = {};
    for (var k in props) o[k] = props[k];
    o.spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
    return o;
}
pstyle("SICoverHead", { fontStyle: "Bold", pointSize: 16, leading: 19.2, justification: C, spaceAfter: 150 });
pstyle("SICoverTitle", { fontStyle: "Bold", pointSize: 16, leading: 19.2, justification: C });
pstyle("SISection",  span({ fontStyle: "Bold", pointSize: 12, leading: 14.4, fillColor: RED,
                            spaceAfter: 6, keepWithNext: 2, startParagraph: StartParagraph.NEXT_PAGE }));
pstyle("SIMovieTitle", span({ fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8, justification: J,
                              keepWithNext: 2, keepAllLinesTogether: true }));
pstyle("SIMovie",    span({ pointSize: 9, leading: 10.8, justification: J, spaceAfter: 8,
                            keepAllLinesTogether: true }));
// A supplementary figure and its caption stay on one page: the figure keeps with the
// caption, and the caption is never split.
pstyle("SIFigure",   span({ justification: C, leading: Leading.AUTO, spaceBefore: 6, spaceAfter: 6,
                            keepWithNext: 1 }));
pstyle("SIFigureNewPage", span({ justification: C, leading: Leading.AUTO, spaceAfter: 6,
                                 keepWithNext: 1, startParagraph: StartParagraph.NEXT_PAGE }));
pstyle("SICaption",  span({ pointSize: 9, leading: 10.8, spaceBefore: 2, spaceAfter: 12,
                            justification: J, keepAllLinesTogether: true }));
pstyle("SICaptionTitle", span({ fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8, justification: J,
                                keepAllLinesTogether: true, keepWithNext: 5 }));
// Tables: caption below the table, as in the reference supplement.
pstyle("SITable",    { leading: Leading.AUTO, spaceBefore: 4, spaceAfter: 4 });
pstyle("SITableWide", span({ leading: Leading.AUTO, spaceBefore: 4, spaceAfter: 4 }));
pstyle("TableCaption", { pointSize: 9, leading: 10.8, spaceAfter: 14, justification: J,
                         keepAllLinesTogether: true });
pstyle("TableCaptionWide", span({ pointSize: 9, leading: 10.8, spaceAfter: 14, justification: J,
                                  keepAllLinesTogether: true }));
pstyle("TableCaptionTitle", { fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8, justification: J,
                              keepWithNext: 5 });
pstyle("TableCaptionTitleWide", span({ fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8,
                                       justification: J, keepWithNext: 5 }));
pstyle("SINote",     { pointSize: 8, leading: 9.6, spaceAfter: 12, justification: J });
pstyle("SINoteWide", span({ pointSize: 8, leading: 9.6, spaceAfter: 12, justification: J }));
pstyle("TableCell",  { pointSize: 7.5, leading: 9 });

with (app.taggedTextImportPreferences) {
    styleConflict = StyleConflict.PUBLICATION_DEFINITION;
    removeTextFormatting = false;
    useTypographersQuotes = true;
}

// ---------------------------------------------------------------------------
// Geometry helpers
// ---------------------------------------------------------------------------
function isRight(page) { return page.side != PageSideOptions.LEFT_HAND; }
function textBox(page) {   // [y1, x1, y2, x2] of the live area
    var l = isRight(page) ? CFG.inside : CFG.outside;
    var r = isRight(page) ? CFG.outside : CFG.inside;
    return [CFG.top, l, CFG.pageH - CFG.bottom, CFG.pageW - r];
}

// ---------------------------------------------------------------------------
// Master pages: footer rule, running title, page number
// ---------------------------------------------------------------------------
var master = doc.masterSpreads[0];
// The running title is shown in full, on up to CFG.footerLines (balanced) lines, clear
// of the page number. Only if it still doesn't fit is it shortened at a word boundary
// and ended with an ellipsis.
var FOOTER_TITLE = null;
function fitFooterTitle(frame, text) {
    var words = text.split(/\s+/), t = text;
    for (;;) {
        frame.parentStory.contents = t;   // the whole story, including any overset text
        frame.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("Footer");
        doc.recompose();
        if (!frame.overflows && frame.lines.length <= CFG.footerLines) return t;
        if (words.length <= 1) return t;
        words.pop();
        t = words.join(" ").replace(/[\s,;:.\-\u2013\u2014]+$/, "") + "\u2026";
    }
}
for (var mp = 0; mp < master.pages.length; mp++) {
    var pg = master.pages[mp], tb = textBox(pg);
    var ln = pg.graphicLines.add();
    ln.geometricBounds = [CFG.footerRule, tb[1], CFG.footerRule, tb[3]];
    ln.strokeWeight = 1; ln.strokeColor = BLACK;
    var rt = pg.textFrames.add({ geometricBounds:   // 40 pt kept clear for the page number
        [CFG.footerTop, tb[1] + 40, CFG.footerTop + 12 * CFG.footerLines + 2, tb[3] - 40] });
    if (FOOTER_TITLE === null) {
        FOOTER_TITLE = fitFooterTitle(rt, M.running_title);
        if (FOOTER_TITLE != M.running_title) LOG.push("Footer title shortened to fit: " + FOOTER_TITLE);
    } else {
        rt.contents = FOOTER_TITLE;
        rt.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("Footer");
    }
    var pn = pg.textFrames.add({ geometricBounds: [CFG.footerTop, tb[1], CFG.footerTop + 14, tb[3]] });
    pn.contents = SpecialCharacters.AUTO_PAGE_NUMBER;
    pn.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("PageNum");
}

// ---------------------------------------------------------------------------
// Page 1: title block and bottom-aligned abstract
// ---------------------------------------------------------------------------
if (!SI) layoutTitlePage();
else layoutCover();

// Supplement cover: "Supplementary Information" above the main paper's title block.
function layoutCover() {
    var p1 = doc.pages[0], b1 = textBox(p1);
    var cf = p1.textFrames.add({ geometricBounds: [120, b1[1], 720, b1[3]] });
    cf.place(File(BUILD + "/cover.txt"));
    doc.recompose();
    if (cf.overflows) LOG.push("WARNING: supplement cover text overflows.");
}

// Page 1. The title block starts 160 pt down with generous gaps; the abstract sits at the
// foot of the page. If it all doesn't fit, tighten step by step: less white space above
// and between the title blocks first, then the abstract one size smaller.
function layoutTitlePage() {
    var p1 = doc.pages[0], b1 = textBox(p1);
    var steps = [
        { top: 160, authors: 54, affil: 56, abs: 11 },
        { top: 110, authors: 36, affil: 36, abs: 11 },
        { top: 72,  authors: 24, affil: 24, abs: 11 },
        { top: 72,  authors: 24, affil: 24, abs: 10 },
        { top: CFG.top, authors: 18, affil: 18, abs: 10 }
    ];
    var tf = null, af = null;
    for (var k = 0; k < steps.length; k++) {
        var st = steps[k];
        doc.paragraphStyles.itemByName("Authors").spaceBefore = st.authors;
        doc.paragraphStyles.itemByName("AffilFirst").spaceBefore = st.affil;
        var abs = doc.paragraphStyles.itemByName("Abstract");
        abs.pointSize = st.abs; abs.leading = st.abs * 1.2;
        if (tf) { tf.remove(); af.remove(); }
        tf = p1.textFrames.add({ geometricBounds: [st.top, b1[1], 480, b1[3]] });
        tf.place(File(BUILD + "/title.txt"));
        doc.recompose();
        var titleBottom = tf.lines.length ? tf.lines[-1].baseline + 6 : 480;
        tf.geometricBounds = [st.top, b1[1], Math.max(titleBottom, st.top + 20), b1[3]];
        af = p1.textFrames.add({ geometricBounds:
            [Math.min(480, titleBottom + 24), b1[1], CFG.abstractBottom, b1[3]] });
        af.place(File(BUILD + "/abstract.txt"));
        af.textFramePreferences.verticalJustification = VerticalJustification.BOTTOM_ALIGN;
        doc.recompose();
        if (!tf.overflows && !af.overflows) {
            if (k > 0) LOG.push("Title page tightened to fit (step " + k + ": title at " + st.top +
                                " pt, abstract " + st.abs + " pt).");
            return;
        }
    }
    LOG.push("WARNING: title page text overflows even when tightened; shorten the abstract.");
}

// ---------------------------------------------------------------------------
// Body: two-column story threaded over as many pages as needed. In the paper each
// column is its own text frame (left -> right -> next page), so each column's height can
// be changed by hand afterwards. The supplement keeps one two-column frame per page,
// because its captions, wide tables and headings span both columns.
// ---------------------------------------------------------------------------
function addBodyPage(prevFrame, page) {
    page = page || doc.pages.add(LocationOptions.AT_END);
    var tb = textBox(page);
    if (SI) {
        var fr = page.textFrames.add({ geometricBounds: tb });
        fr.textFramePreferences.textColumnCount = 2;
        fr.textFramePreferences.textColumnGutter = CFG.gutter;
        if (prevFrame) prevFrame.nextTextFrame = fr;
        return fr;
    }
    var colW = (tb[3] - tb[1] - CFG.gutter) / 2;
    var left = page.textFrames.add({ geometricBounds: [tb[0], tb[1], tb[2], tb[1] + colW] });
    var right = page.textFrames.add({ geometricBounds: [tb[0], tb[3] - colW, tb[2], tb[3]] });
    if (prevFrame) prevFrame.nextTextFrame = left;
    left.nextTextFrame = right;
    return right;   // the last frame of the thread
}
var first = addBodyPage(null);   // the body (or supplement content) starts on page 2
if (!SI) first = first.previousTextFrame;   // the page's left column
first.place(File(BUILD + (SI ? "/si.txt" : "/body.txt")));
var story = first.parentStory;

function flow() {
    doc.recompose();
    var last = story.textContainers[story.textContainers.length - 1];
    var guard = 0;
    while (last.overflows && guard++ < 200) { last = addBodyPage(last); doc.recompose(); }
}
flow();

// Equations typeset with TeX (equations.pdf, one page each) replace their [[EQ:k]]
// placeholders as inline graphics sitting on the text baseline.
function placeEquations() {
    var eqs = M.equations || [], n = 0;
    if (!eqs.length) return;
    app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_MEDIA;   // keep TeX's exact box
    var maxW = COL_W - 2 * 28;                               // room for the number
    for (var k = 0; k < eqs.length; k++) {
        var e = eqs[k];
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "\\[\\[EQ:" + k + "\\]\\]";
        var hits = story.findGrep();
        if (!hits.length || !e.typeset) continue;
        var hit = hits[0], para = hit.paragraphs[0], idx = hit.index;
        hit.remove();
        app.pdfPlacePreferences.pageNumber = e.page;
        var g = story.insertionPoints[idx].place(File(e.file))[0], rect = g.parent;
        var sc = Math.min(1, maxW / e.width);
        if (sc < 1) {
            g.absoluteHorizontalScale = g.absoluteVerticalScale = sc * 100;
            rect.fit(FitOptions.FRAME_TO_CONTENT);
            LOG.push("Equation " + (e.number || k + 1) + " scaled to " + Math.round(sc * 100) + "% to fit the column.");
        }
        rect.strokeWeight = 0; rect.strokeColor = doc.swatches.itemByName("None");
        story.characters[idx].baselineShift = -e.depth * sc;   // TeX baseline on the text baseline
        para.leading = Math.max(12, e.height * sc + 3);
        para.spaceAfter = CFG.paragraphSpacing + Math.max(0, e.depth * sc - 3);
        n++;
    }
    app.findGrepPreferences = NothingEnum.NOTHING;
    if (n) LOG.push("Equations: " + n + " typeset with TeX in " + CFG.font + ".");
}
placeEquations();
flow();

// ---------------------------------------------------------------------------
// Figures
// ---------------------------------------------------------------------------
// Index (in story.contents) of the first in-text citation of main Figure n.
function firstCitation(n) {
    var s = story.contents;
    var re = /Fig(?:ure)?s?\.?\s+((?:S?\d+[A-Za-z]*(?:\s*[-–,]\s*|\s+and\s+)?)+)/g, m;
    while ((m = re.exec(s)) !== null) {
        // "Supplementary Figure 2" / "Extended Data Fig. 2" are not main figures
        if (/(Supplementa(l|ry)|Extended Data|Appendix)\s*$/i.test(s.substring(m.index - 20, m.index))) continue;
        var nums = m[1].match(/S?\d+/g);
        for (var i = 0; i < nums.length; i++)
            if (nums[i] == String(n)) return m.index;
    }
    return -1;
}

// First main-heading paragraph (e.g. DISCUSSION) after story position `at`, or null.
function nextHeading(at) {
    var ps = story.paragraphs.everyItem().getElements();
    for (var i = 0; i < ps.length; i++) {
        var nm = ps[i].appliedParagraphStyle.name;
        if ((nm == "H1" || nm == "H1Back") && ps[i].index > at) return ps[i];
    }
    return null;
}
function pageOf(text) {
    var tf = text.parentTextFrames;
    return tf.length ? tf[0].parentPage.documentOffset : -1;
}

var slots = {}; // page index -> {top: usedHeight, bottom: usedHeight}
function slot(i) { if (!slots[i]) slots[i] = { top: 0, bottom: 0 }; return slots[i]; }

// Figure frames have no outline unless asked for (`mudbrick docx2pdf --figure-border`).
// Order matters: InDesign restores a 1 pt stroke if the weight is set after the colour
// is None, so set the weight first.
function figureStroke(frame) {
    if (M.figure_border) {
        frame.strokeWeight = 0.5; frame.strokeColor = BLACK;
        frame.strokeAlignment = StrokeAlignment.INSIDE_ALIGNMENT;
    } else {
        frame.strokeWeight = 0; frame.strokeColor = doc.swatches.itemByName("None");
    }
}

function rule(page, y, x1, x2) {
    var l = page.graphicLines.add();
    l.geometricBounds = [y, x1, y, x2]; l.strokeWeight = 1; l.strokeColor = BLACK;
    return l;
}

// Shrink the body-text frames (both columns) on page pi to the space figures leave free.
// (Text wrap is avoided: InDesign drops some placed AI/PDF art under a wrap on export.)
function reserve(pi) {
    var tb = textBox(doc.pages[pi]), sl = slot(pi), frs = [], tc = story.textContainers;
    for (var i = 0; i < tc.length; i++)
        if (tc[i].parentPage && tc[i].parentPage.documentOffset == pi) frs.push(tc[i]);
    var y1 = tb[0] + sl.top, y2 = tb[2] - sl.bottom;
    for (var k = frs.length - 1; k >= 0; k--) {
        var b = frs[k].geometricBounds;
        if (y2 - y1 < 3 * 12) frs[k].remove();  // thread skips straight to the next page
        else frs[k].geometricBounds = [y1, b[1], y2, b[3]];
    }
}

// Text frame holding a legend.
function legendFrame(page, tb, height, file) {
    var f = page.textFrames.add({ geometricBounds: [tb[0], tb[1], tb[0] + height, tb[3]] });
    f.textFramePreferences.ignoreWrap = true;
    f.place(file);
    doc.recompose();
    return f;
}

// Height actually used by the text in a frame.
function usedHeight(f) {
    var ls = f.lines;
    if (!ls.length) return 0;
    return ls[ls.length - 1].baseline - f.geometricBounds[0] + 3;
}

// Put a frame's placed artwork so its top-left corner sits (ox, oy) from the frame's
// top-left corner (negative = the artwork starts above/left of the frame, i.e. it is
// trimmed). Never set a graphic's geometricBounds directly: with page rulers InDesign
// reads and writes a placed graphic's bounds relative to the spread, not its page, so
// on a right-hand page the artwork lands on the facing page, under that page's text,
// while the empty frame stays put.
function positionGraphic(rect, ox, oy) {
    var g = rect.allGraphics[0];
    var where = [AnchorPoint.TOP_LEFT_ANCHOR, BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS];
    var f = rect.resolve(where, CoordinateSpaces.PASTEBOARD_COORDINATES)[0];
    var c = g.resolve(where, CoordinateSpaces.PASTEBOARD_COORDINATES)[0];
    g.move(undefined, [f[0] + ox - c[0], f[1] + oy - c[1]]);
    return g;
}

function placeFigure(fig, forcePi) {
    var file = fig.file ? File(fig.file) : null;
    var crop = fig.crop || [0, 0, 1, 1];
    var at = firstCitation(fig.num);
    var citePage = 1;
    if (at >= 0) citePage = story.characters[at].parentTextFrames[0].parentPage.documentOffset;
    // A figure goes at the top of the page right after its first citation (never
    // before it). The next section (e.g. DISCUSSION) may start on the citation page;
    // with CFG.sectionAfterFigures it is pushed to a new page below the figure instead.
    var pi = forcePi || Math.max(citePage + 1, 1);
    var head = at >= 0 ? nextHeading(at) : null;
    if (!forcePi && CFG.sectionAfterFigures && head && pageOf(head) >= 1 && pageOf(head) < pi) {
        head.startParagraph = StartParagraph.NEXT_PAGE;
        flow();
        LOG.push(head.contents.replace(/\s+$/, "") + " moved to a new page so Figure " + fig.num +
                 " follows its citation and precedes that section.");
    }

    while (pi >= doc.pages.length) addBodyPage(story.textContainers[story.textContainers.length - 1]);
    var page = doc.pages[pi], tb = textBox(page), W = tb[3] - tb[1], avail = tb[2] - tb[0];
    var legendFile = File(BUILD + "/" + fig.legend), lineH = 10.8;

    // Measure the full legend.
    var probe = legendFrame(page, tb, 5000, legendFile);
    var legendH = usedHeight(probe);
    probe.remove();

    // Figure at its own resolution, cropped to its content, limited to the text width.
    var rect = page.rectangles.add({ geometricBounds: [tb[0], tb[1], tb[0] + 100, tb[1] + 100] });
    figureStroke(rect);
    var fullW = W, fullH = W * 0.6;
    if (file && file.exists) {
        // PDF/AI figures are placed by their crop box (the Illustrator artboard),
        // the same area the whitespace trim was measured on.
        app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_PDF;
        app.pdfPlacePreferences.transparentBackground = true;
        rect.place(file);
        var gb = rect.allGraphics[0].geometricBounds;
        fullW = gb[3] - gb[1]; fullH = gb[2] - gb[0];
    } else {
        file = null; crop = [0, 0, 1, 1];
        LOG.push("WARNING: no file found for Figure " + fig.num + " (placeholder used).");
        rect.strokeWeight = 0.5;
    }
    var natW = fullW * (crop[2] - crop[0]), natH = fullH * (crop[3] - crop[1]);
    // Never resized: scaling the artwork would change its font sizes. Artwork wider
    // than the text column is centred and extends into the side margins.
    var s = 1, wideWarning = null;
    if (file && natW > W + 0.5) {
        var side = (natW - W) / 2;
        wideWarning = ("WARNING: Figure " + fig.num + " artwork is " + Math.round(natW) + " pt wide, " +
                 Math.round(natW - W) + " pt wider than the " + Math.round(W) + "-pt text column; " +
                 "it overprints " + Math.round(side) + " pt into each side margin" +
                 (side > Math.min(tb[1], CFG.pageW - tb[3]) ? " and runs off the page edge" : "") + ".");
    }

    // Decide how big it is and how much legend fits with it.
    var extras = CFG.figGap + CFG.ruleGap + CFG.wrapGap;
    var minLegend = Math.min(legendH, CFG.minLegendLines * lineH);
    var room = avail - slot(pi).top - slot(pi).bottom - extras;
    if (room < natH * s * 0.6 + minLegend && (slot(pi).top || slot(pi).bottom)) {
        rect.remove(); return placeFigure(fig, pi + 1);
    }
    var figW = natW * s, figH = natH * s;
    var legendRoom = room - figH;
    if (legendRoom < lineH * 1.5) {
        LOG.push("WARNING: Figure " + fig.num + " artwork is " + Math.round(natH) + " pt tall and leaves " +
                 "no room for its legend on page " + (pi + 1) + "; it overprints the page foot.");
        legendRoom = lineH * 1.5;
    }

    // Set the legend; panels that do not fit continue on the next page.
    var lf = legendFrame(page, tb, legendRoom, legendFile);
    var split = lf.overflows;
    if (split) {   // cut at the last line that fits; the rest continues on the next page
        var keep = lf.characters.length;
        continueLegend(fig, keep, pi + 1);
        var lst = lf.parentStory;
        lst.characters.itemByRange(keep, lst.characters.length - 1).remove();
        var lc = lst.characters;
        if (lc.length && lc[-1].contents == "\r") lc[-1].remove();
        doc.recompose();
    }
    var legendHere = usedHeight(lf);

    var atTop = CFG.figurePosition == "top" ? !slot(pi).top : !!slot(pi).bottom;
    var blockH = figH + CFG.figGap + legendHere;
    var y0 = atTop ? tb[0] + slot(pi).top : tb[2] - slot(pi).bottom - blockH;
    var x0 = tb[1] + (W - figW) / 2;

    rect.geometricBounds = [y0, x0, y0 + figH, x0 + figW];
    var ox = -crop[0] * fullW, oy = -crop[1] * fullH;   // at 100 %: shift so the content fills the frame
    if (file) positionGraphic(rect, ox, oy);
    lf.geometricBounds = [y0 + figH + CFG.figGap, tb[1], y0 + blockH, tb[3]];
    if (file) PLACED.push({ page: page, file: file, frame: rect.geometricBounds, ox: ox, oy: oy, rect: rect });

    var fullPage = (avail - slot(pi).top - slot(pi).bottom - blockH - extras) < 3 * 12;
    if (atTop) {
        var ry = y0 + blockH + CFG.ruleGap;
        if (!fullPage) rule(page, ry, tb[1], tb[3]);
        slot(pi).top += fullPage ? avail : ry + CFG.wrapGap - y0;
    } else {
        var ry2 = y0 - CFG.ruleGap;
        if (!fullPage) rule(page, ry2, tb[1], tb[3]);
        slot(pi).bottom += fullPage ? avail : tb[2] - (ry2 - CFG.wrapGap);
    }
    reserve(pi);

    flow();
    if (wideWarning) LOG.push(wideWarning);
    LOG.push("Figure " + fig.num + ": cited p" + (citePage + 1) + ", placed p" + (pi + 1) +
             (atTop ? " top" : " bottom") + ", " + Math.round(100 * s) + "% of file size" +
             (split ? ", legend continued on p" + (pi + 2) : ""));
}

// Legend text from character `from` on continues at the foot of page pi.
function continueLegend(fig, from, pi) {
    while (pi >= doc.pages.length) addBodyPage(story.textContainers[story.textContainers.length - 1]);
    var page = doc.pages[pi], tb = textBox(page);
    var cf = legendFrame(page, tb, tb[2] - tb[0], File(BUILD + "/" + fig.legend));
    var st = cf.parentStory;
    st.characters.itemByRange(0, from - 1).remove();
    while (st.characters.length && /^\s$/.test(st.characters[0].contents)) st.characters[0].remove();
    var label = "Figure " + fig.num + " ", note = "(continued from prior page).";
    st.insertionPoints[0].contents = label + note + "\r";
    st.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("LegendCont");
    cstyle("cs_sb_i");
    st.paragraphs[0].characters.itemByRange(label.length, label.length + note.length - 1)
        .appliedCharacterStyle = doc.characterStyles.itemByName("cs_sb_i");
    doc.recompose();
    var h = usedHeight(cf);
    var y0 = tb[2] - slot(pi).bottom - h;
    cf.geometricBounds = [y0, tb[1], tb[2] - slot(pi).bottom, tb[3]];
    var ry = y0 - CFG.ruleGap;
    rule(page, ry, tb[1], tb[3]);
    slot(pi).bottom += tb[2] - (ry - CFG.wrapGap);
    reserve(pi);
}

var PLACED = [];
M.figures.sort(function (a, b) { return a.num - b.num; });
if (SI) buildSupplement();
else for (var fi = 0; fi < M.figures.length; fi++) placeFigure(M.figures[fi]);

// Re-place every figure into a fresh frame with the final geometry. Frames that went
// through the sizing steps above can be silently dropped from the PDF export for some
// Illustrator files (an InDesign quirk); fresh frames export fine.
for (var pf = 0; pf < PLACED.length; pf++) {
    var P = PLACED[pf];
    P.rect.remove();
    var nr = P.page.rectangles.add({ geometricBounds: P.frame });
    figureStroke(nr);
    app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_PDF;
    nr.place(P.file);
    positionGraphic(nr, P.ox, P.oy);
}

// ---------------------------------------------------------------------------
// Supplement: replace [[FIG:n]] / [[TABLE:n]] placeholders with pictures and tables
// ---------------------------------------------------------------------------
function buildSupplement() {
    SI_PLACED = [];
    var tb = textBox(doc.pages[0]), W = tb[3] - tb[1], CW = (W - CFG.gutter) / 2;
    var maxH = tb[2] - tb[0];   // a whole page; siFigure subtracts the caption's height
    for (var guard = 0; guard < 500; guard++) {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "\\[\\[(FIG|TABLE):\\d+\\]\\]";
        var hits = story.findGrep();
        if (!hits.length) break;
        var hit = hits[0], m = hit.contents.match(/(FIG|TABLE):(\d+)/);
        var para = hit.paragraphs[0];
        hit.remove();
        var ip = para.insertionPoints[0];
        if (m[1] == "FIG") siFigure(ip, M.si_figures[+m[2]], W, maxH);
        else {
            var t = M.tables[+m[2]];
            siTable(ip, t, t.wide ? W : CW);
            flow();
            keepTableWithCaption(para, t.wide, CW);
        }
        flow();
    }
    app.findGrepPreferences = NothingEnum.NOTHING;
    placeSIFigures();
    if (START_PAGE > 0)   // continue the main paper's page numbers
        doc.sections[0].properties = { continueNumbering: false, pageNumberStart: START_PAGE,
                                       includeSectionPrefix: false };
    else
        doc.sections[0].properties = { sectionPrefix: "S", includeSectionPrefix: true };
    LOG.push("Supplement: " + M.si_figures.length + " figures, " + M.tables.length + " tables");
}

// Height the caption paragraphs after `para` (up to the next figure/table) take at width W.
function captionHeight(para, W) {
    var tmp = doc.pages[0].textFrames.add({ geometricBounds: [0, -W - 100, 5000, -100] });
    var p = para, h = 0;
    for (var guard = 0; guard < 4; guard++) {
        p = story.paragraphs.nextItem(p);
        if (!p || !p.isValid) break;
        var nm = p.appliedParagraphStyle.name;
        if (nm != "SICaption" && nm != "SICaptionTitle") break;
        p.duplicate(LocationOptions.AT_END, tmp.parentStory.insertionPoints[-1]);
        h += p.spaceBefore;
    }
    doc.recompose();
    var ls = tmp.lines;
    if (ls.length) h += ls[-1].baseline - tmp.geometricBounds[0] + 4;
    tmp.remove();
    return h;
}

// Full caption width if figure + caption still fit on one page; otherwise as large as
// fits. (Pictures may be enlarged to reach the caption width.)
// The caption sits below its table. If it ends up in a later column or page, the
// table moves there too, so they stay together.
function keepTableWithCaption(tablePara, wide, CW) {
    var cap = story.paragraphs.nextItem(tablePara);
    if (!cap || !cap.isValid || !/^TableCaption/.test(cap.appliedParagraphStyle.name)) return;
    var pt = pageOf(tablePara), pc = pageOf(cap);
    var dx = Math.abs(cap.insertionPoints[0].horizontalOffset -
                      tablePara.insertionPoints[0].horizontalOffset);
    if (pt >= 0 && pc >= 0 && (pc > pt || (!wide && dx > CW / 2))) {
        tablePara.startParagraph = wide ? StartParagraph.NEXT_FRAME : StartParagraph.NEXT_COLUMN;
        flow();
    }
}

// Supplementary figures are never resized (that would change their font sizes). The
// text only reserves their space with an empty inline frame; once the layout is final,
// placeSIFigures() puts each figure at 100 % on its page, centred on the text column over
// that space. A figure larger than the space overprints the margins or its caption.
// (Placing them inline instead makes InDesign drop the artwork of a figure whose
// paragraph is later pushed to a new page, and cannot let a figure overhang the text.)
var SI_PLACED = [];   // reset in buildSupplement(), which runs before this line

function siFigure(ip, fig, W, maxH) {
    app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_PDF;
    app.pdfPlacePreferences.transparentBackground = true;
    var para = ip.paragraphs[0];
    var capH = captionHeight(para, W);
    // natural size of the artwork, trimmed to its content
    var tmp = doc.pages[0].rectangles.add({ geometricBounds: [0, 0, 10, 10] });
    var gb = tmp.place(File(fig.file))[0].geometricBounds;
    tmp.remove();
    var crop = fig.crop || [0, 0, 1, 1];
    var fullW = gb[3] - gb[1], fullH = gb[2] - gb[0];
    var natW = fullW * (crop[2] - crop[0]), natH = fullH * (crop[3] - crop[1]);
    var maxRoom = Math.max(72, maxH - capH - 20);   // most a page can hold above the caption
    var spacer = ip.rectangles.add({ geometricBounds: [0, 0, 10, 10] });
    var rb = spacer.geometricBounds;
    spacer.geometricBounds = [rb[0], rb[1], rb[0] + Math.min(natH, maxRoom), rb[1] + Math.min(natW, W)];
    spacer.strokeWeight = 0;
    spacer.strokeColor = doc.swatches.itemByName("None");
    spacer.fillColor = doc.swatches.itemByName("None");
    SI_PLACED.push({ spacer: spacer, fig: fig, crop: crop, fullW: fullW, fullH: fullH,
                     natW: natW, natH: natH, W: W });
}

function placeSIFigures() {
    for (var i = 0; i < SI_PLACED.length; i++) {
        var it = SI_PLACED[i], sp = it.spacer, page = sp.parentPage, n = it.fig.num;
        if (!page) { LOG.push("WARNING: Supplementary Figure " + n + " did not fit on any page."); continue; }
        var sb = sp.geometricBounds, tb = textBox(page), W = tb[3] - tb[1];
        var x0 = tb[1] + (W - it.natW) / 2;
        var over = it.natH - (sb[2] - sb[0]);
        // a figure taller than its space first uses the empty top margin, then runs down
        var up = over > 0 ? Math.min(over, Math.max(0, sb[0] - 18)) : 0;
        var y0 = sb[0] - up;
        var rect = page.rectangles.add({ geometricBounds: [y0, x0, y0 + it.natH, x0 + it.natW] });
        rect.place(File(it.fig.file));
        var ox = -it.crop[0] * it.fullW, oy = -it.crop[1] * it.fullH;
        positionGraphic(rect, ox, oy);
        // the final frame comes from the re-place pass after the layout, as for main figures
        PLACED.push({ page: page, file: File(it.fig.file), frame: rect.geometricBounds, ox: ox, oy: oy, rect: rect });
        figureStroke(rect);
        if (it.natW > W + 0.5) {
            var side = (it.natW - W) / 2;
            LOG.push("WARNING: Supplementary Figure " + n + " artwork is " + Math.round(it.natW) +
                     " pt wide, " + Math.round(it.natW - W) + " pt wider than the " + Math.round(W) +
                     "-pt text column; it overprints " + Math.round(side) + " pt into each side margin" +
                     (side > Math.min(tb[1], CFG.pageW - tb[3]) ? " and runs off the page edge" : "") + ".");
        }
        if (over > 0.5) {
            var down = over - up;
            LOG.push("WARNING: Supplementary Figure " + n + " artwork is " + Math.round(it.natH) +
                     " pt tall, " + Math.round(over) + " pt more than fits with its caption on page " +
                     page.name + "; it overprints " +
                     (up > 0.5 ? Math.round(up) + " pt of the top margin" : "") +
                     (up > 0.5 && down > 0.5 ? " and " : "") +
                     (down > 0.5 ? Math.round(down) + " pt of its caption" +
                      (y0 + it.natH > CFG.footerRule ? " and the page footer" : "") : "") + ".");
        }
    }
}

function siTable(ip, t, W) {
    var nrows = t.rows.length, ncols = t.grid.length, total = 0, r, c, i;
    for (i = 0; i < ncols; i++) total += t.grid[i];
    var tbl = ip.tables.add({ bodyRowCount: nrows, columnCount: ncols });
    tbl.width = W;
    for (i = 0; i < ncols; i++) tbl.columns[i].width = W * t.grid[i] / total;
    var cellStyle = doc.paragraphStyles.itemByName("TableCell"), merges = [];
    for (r = 0; r < nrows; r++) {
        var col = 0, row = t.rows[r];
        for (c = 0; c < row.length && col < ncols; c++) {
            var cell = tbl.rows[r].cells[col];
            cell.contents = row[c].t;
            cell.texts[0].appliedParagraphStyle = cellStyle;
            if (row[c].b) cell.texts[0].fontStyle = "Bold";
            if (row[c].span > 1) merges.push([r, col, Math.min(ncols, col + row[c].span) - 1]);
            col += row[c].span;
        }
        for (; col < ncols; col++) tbl.rows[r].cells[col].texts[0].appliedParagraphStyle = cellStyle;
    }
    for (i = merges.length - 1; i >= 0; i--)   // right-to-left, bottom-up keeps indices valid
        tbl.rows[merges[i][0]].cells[merges[i][1]].merge(tbl.rows[merges[i][0]].cells[merges[i][2]]);
    // Minimal style: a rule at the top and bottom of the table and under each row of
    // column headings; section rows that span the table get a light band. No other lines.
    var allBold = true;
    for (c = 0; c < t.rows[0].length; c++) if (!t.rows[0][c].b && t.rows[0][c].t) allBold = false;
    // repeat the first row when the table breaks (convert it; setting headerRowCount
    // would insert a new, empty header row instead)
    if (allBold && nrows > 2) tbl.rows[0].rowType = RowTypes.HEADER_ROW;
    var cells = tbl.cells.everyItem();
    cells.topInset = cells.bottomInset = 2;
    cells.leftInset = cells.rightInset = 3;
    cells.topEdgeStrokeWeight = cells.bottomEdgeStrokeWeight = 0;
    cells.leftEdgeStrokeWeight = cells.rightEdgeStrokeWeight = 0;
    for (r = 0; r < nrows; r++) {
        var row = t.rows[r], rowCells = tbl.rows[r].cells.everyItem(), bold = true, any = false;
        for (c = 0; c < row.length; c++) { if (row[c].t) { any = true; if (!row[c].b) bold = false; } }
        var band = row.length == 1 && row[0].span >= ncols;
        if (band && bold && any) {            // section row, e.g. "Overall performance"
            rowCells.fillColor = BLACK; rowCells.fillTint = 12;
            rowCells.topInset = rowCells.bottomInset = 3;
        } else if (bold && any && row.length > 1) {   // column headings
            rowCells.bottomEdgeStrokeWeight = 0.5;
            rowCells.bottomEdgeStrokeColor = BLACK; rowCells.bottomEdgeStrokeTint = 100;
        }
    }
    var top = tbl.rows[0].cells.everyItem(), bottom = tbl.rows[nrows - 1].cells.everyItem();
    top.topEdgeStrokeWeight = 0.75; top.topEdgeStrokeColor = BLACK; top.topEdgeStrokeTint = 100;
    bottom.bottomEdgeStrokeWeight = 0.75; bottom.bottomEdgeStrokeColor = BLACK;
    bottom.bottomEdgeStrokeTint = 100;

    // A table that fits on one page is never split: chain its rows together so it moves
    // (with its caption, which keeps with the table) to the next page when space runs
    // out. Taller tables break normally and repeat their header row.
    doc.recompose();
    var tb = textBox(doc.pages[0]), h = 0;
    for (r = 0; r < tbl.rows.length; r++) h += tbl.rows[r].height;
    if (h < (tb[2] - tb[0]) - 60)
        for (r = 0; r < tbl.rows.length - 1; r++)   // header rows travel with the table anyway
            if (tbl.rows[r].rowType == RowTypes.BODY_ROW) tbl.rows[r].keepWithNextRow = true;
}

// ---------------------------------------------------------------------------
// Clean up and export
// ---------------------------------------------------------------------------
flow();
// drop trailing empty pages
for (var pi2 = doc.pages.length - 1; pi2 > 1; pi2--) {
    var pgx = doc.pages[pi2], tfs = pgx.textFrames, empty = pgx.allPageItems.length <= 1;
    if (empty && tfs.length == 1 && tfs[0].characters.length == 0) pgx.remove(); else break;
}
doc.recompose();

// Columns end flush at the bottom. When a heading moves to the next column together
// with its first lines (leaving more than a line empty), the gap is shared out between
// that column's ordinary paragraphs only, so headings keep exactly their usual spacing.
// Columns with the normal under-one-line remainder are left alone.
var FILLABLE = { "Body": 1, "BodyTight": 1, "Equation": 1, "EquationTeX": 1 };   // not references: their spacing stays fixed
function fillColumnGaps() {
    var conts = story.textContainers, filled = 0;
    for (var k = 0; k < conts.length - 1; k++) {          // the story's last frame stays as is
        var fr = conts[k], bottom = fr.geometricBounds[2], cols = fr.textColumns;
        for (var c = 0; c < cols.length; c++) {
            var ls = cols[c].lines;
            if (!ls.length) continue;
            var gap = bottom - (ls[-1].baseline + 3);
            if (gap <= 13) continue;                       // normal remainder (< one line)
            var ps = cols[c].paragraphs, slots = [];
            for (var i = 0; i < ps.length - 1; i++) {
                var a = ps[i].appliedParagraphStyle.name, b = ps[i + 1].appliedParagraphStyle.name;
                if (FILLABLE[a] && FILLABLE[b]) slots.push(ps[i]);
            }
            if (!slots.length) continue;
            var extra = Math.min(CFG.columnFillLimit, (gap - 2) / slots.length);
            for (var j = 0; j < slots.length; j++) slots[j].spaceAfter = slots[j].spaceAfter + extra;
            doc.recompose();
            filled++;
        }
    }
    return filled;
}
if (CFG.fillColumns && !SI) {
    var nfilled = fillColumnGaps();
    if (nfilled) LOG.push("Filled " + nfilled + " column gap(s) left by headings moving to the next column.");
}

// Print-quality PDF: images kept at full resolution with lossless compression,
// colours left unconverted, fonts embedded in full, transparency preserved.
function printPreset() {
    var name = "Preprint - Maximum Quality";
    var p = app.pdfExportPresets.itemByName(name);
    if (!p.isValid) p = app.pdfExportPresets.add({ name: name });
    p.properties = {
        acrobatCompatibility: AcrobatCompatibility.ACROBAT_7,   // PDF 1.6, live transparency
        standardsCompliance: PDFXStandards.NONE,
        pdfColorSpace: PDFColorSpace.UNCHANGED_COLOR_SPACE,
        // Images far finer than any printer resolves (e.g. 10,000-px maps embedded in
        // Illustrator files) are reduced to CFG.printPPI; everything else is untouched.
        colorBitmapSampling: Sampling.BICUBIC_DOWNSAMPLE,
        colorBitmapSamplingDPI: CFG.printPPI,
        thresholdToCompressColor: CFG.downsampleAbovePPI,
        grayscaleBitmapSampling: Sampling.BICUBIC_DOWNSAMPLE,
        grayscaleBitmapSamplingDPI: CFG.printPPI,
        thresholdToCompressGray: CFG.downsampleAbovePPI,
        monochromeBitmapSampling: Sampling.BICUBIC_DOWNSAMPLE,
        monochromeBitmapSamplingDPI: 1200,
        thresholdToCompressMonochrome: 1800,
        colorBitmapCompression: BitmapCompression.ZIP,
        grayscaleBitmapCompression: BitmapCompression.ZIP,
        monochromeBitmapCompression: MonoBitmapCompression.ZIP,
        compressTextAndLineArt: true,
        cropImagesToFrames: true,
        subsetFontsBelow: 0,                                     // embed complete fonts
        includeICCProfiles: ICCProfiles.INCLUDE_ALL,
        exportReaderSpreads: false,
        generateThumbnails: false,
        optimizePDF: false,
        cropMarks: false, bleedMarks: false, registrationMarks: false,
        colorBars: false, pageInformationMarks: false
    };
    return p;
}
// InDesign cannot resample some images embedded in Illustrator files (soft-masked
// images fail with "Unable to read TIFF"), so fall back to gentler settings.
var preset = printPreset(), ladder = [[CFG.printPPI, CFG.downsampleAbovePPI], [600, 900], null];
for (var li = 0; li < ladder.length; li++) {
    var L = ladder[li];
    if (L) {
        preset.colorBitmapSampling = preset.grayscaleBitmapSampling = Sampling.BICUBIC_DOWNSAMPLE;
        preset.colorBitmapSamplingDPI = preset.grayscaleBitmapSamplingDPI = L[0];
        preset.thresholdToCompressColor = preset.thresholdToCompressGray = L[1];
    } else {
        preset.colorBitmapSampling = preset.grayscaleBitmapSampling = Sampling.NONE;
    }
    try {
        doc.exportFile(ExportFormat.PDF_TYPE, OUTPDF, false, preset);
        LOG.push("PDF images: " + (L ? "above " + L[1] + " ppi reduced to " + L[0] + " ppi" : "full resolution (no downsampling)"));
        break;
    } catch (e) {
        LOG.push("Export with " + (L ? L[0] + "/" + L[1] + " ppi" : "no downsampling") + " failed (" + e.message + "); trying gentler settings.");
        if (!L) throw e;
    }
}
if (CFG.embedFigures && M.embed_figures !== false) {
    var embedded = 0, failed = [];
    for (var k = doc.links.length - 1; k >= 0; k--) {
        var lk = doc.links[k];
        if (lk.status == LinkStatus.LINK_EMBEDDED) continue;
        try { lk.unlink(); embedded++; } catch (e) { failed.push(lk.name + " (" + e.message + ")"); }
    }
    LOG.push("Embedded " + embedded + " figure file(s) in the .indd" +
             (failed.length ? "; could not embed: " + failed.join(", ") : ""));
}
if (M.embed_figures === false) LOG.push("Figures linked, not embedded: keep the figure folder next to the .indd");
doc.save(File(OUTPDF.fsName.replace(/\.pdf$/i, "") + ".indd"));
LOG.push("pages: " + doc.pages.length);
doc.close(SaveOptions.NO);
LOG.join("\n");
