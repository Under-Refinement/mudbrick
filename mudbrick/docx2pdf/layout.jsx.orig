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
    footerRule: 741, footerTop: 742,
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
    minFigScale: 0.9,          // figures may shrink to this to keep more legend on the page
    figurePosition: "top",     // "top" or "bottom" of the page after first citation
    abstractBottom: 730,       // foot of the abstract block on page 1 (footer rule at 741)
    sectionAfterFigures: false, // true: push a section heading (e.g. DISCUSSION) to a new
                                // page rather than let it start before a figure cited
                                // at the end of the previous section
    printPPI: 450,             // export: images finer than downsampleAbovePPI are
    downsampleAbovePPI: 600    //   reduced to printPPI (lossless ZIP either way)
};

var BUILD, OUTPDF;
if (typeof arguments !== "undefined" && arguments.length >= 2) {
    BUILD = Folder(arguments[0]); OUTPDF = File(arguments[1]);
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
pstyle("Body",       { justification: J, spaceAfter: 12 });
pstyle("BodyTight",  { justification: J });
pstyle("H1",         { fontStyle: "Bold", pointSize: 12, leading: 14.4, fillColor: RED, keepWithNext: 2, keepAllLinesTogether: true });
pstyle("H1Back",     { fontStyle: "Bold", pointSize: 12, leading: 14.4, fillColor: RED, keepWithNext: 2, keepAllLinesTogether: true });
pstyle("H2",         { fontStyle: "Bold", keepWithNext: 2, keepAllLinesTogether: true });
pstyle("Reference",  { justification: J, pointSize: 8, leading: 9.6, spaceAfter: 6.4,
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
pstyle("LegendTitle", { fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8 });
pstyle("Legend",     { pointSize: 9, leading: 10.8, spaceBefore: 2, keepAllLinesTogether: true });
pstyle("LegendCont", { fontStyle: SEMIBOLD, pointSize: 9, leading: 10.8 });
pstyle("Footer",     { fontStyle: "Italic", justification: C });
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
for (var mp = 0; mp < master.pages.length; mp++) {
    var pg = master.pages[mp], tb = textBox(pg);
    var ln = pg.graphicLines.add();
    ln.geometricBounds = [CFG.footerRule, tb[1], CFG.footerRule, tb[3]];
    ln.strokeWeight = 1; ln.strokeColor = BLACK;
    var rt = pg.textFrames.add({ geometricBounds: [CFG.footerTop, tb[1], CFG.footerTop + 14, tb[3]] });
    rt.contents = M.running_title;
    rt.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("Footer");
    var pn = pg.textFrames.add({ geometricBounds: [CFG.footerTop, tb[1], CFG.footerTop + 14, tb[3]] });
    pn.contents = SpecialCharacters.AUTO_PAGE_NUMBER;
    pn.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("PageNum");
}

// ---------------------------------------------------------------------------
// Page 1: title block and bottom-aligned abstract
// ---------------------------------------------------------------------------
var p1 = doc.pages[0], b1 = textBox(p1);
var tf = p1.textFrames.add({ geometricBounds: [160, b1[1], 480, b1[3]] });
tf.place(File(BUILD + "/title.txt"));
// Abstract sits at the foot of page 1; a long one may use everything below the title block.
doc.recompose();
var titleBottom = tf.lines.length ? tf.lines[-1].baseline + 6 : 480;
var af = p1.textFrames.add({ geometricBounds: [Math.min(480, titleBottom + 24), b1[1], CFG.abstractBottom, b1[3]] });
af.place(File(BUILD + "/abstract.txt"));
af.textFramePreferences.verticalJustification = VerticalJustification.BOTTOM_ALIGN;
doc.recompose();
if (af.overflows) {   // still too long: set the abstract one size smaller
    var abs = doc.paragraphStyles.itemByName("Abstract");
    abs.pointSize = 10; abs.leading = 12;
    doc.recompose();
    LOG.push("Abstract too long for 11 pt on the title page; set at 10/12 pt.");
}
if (tf.overflows || af.overflows) LOG.push("WARNING: title page text overflows.");

// ---------------------------------------------------------------------------
// Body: two-column story threaded over as many pages as needed
// ---------------------------------------------------------------------------
function addBodyPage(prevFrame) {
    var page = doc.pages.add(LocationOptions.AT_END);
    var fr = page.textFrames.add({ geometricBounds: textBox(page) });
    fr.textFramePreferences.textColumnCount = 2;
    fr.textFramePreferences.textColumnGutter = CFG.gutter;
    if (prevFrame) prevFrame.nextTextFrame = fr;
    return fr;
}
var first = addBodyPage(null);
first.place(File(BUILD + "/body.txt"));
var story = first.parentStory;

function flow() {
    doc.recompose();
    var last = story.textContainers[story.textContainers.length - 1];
    var guard = 0;
    while (last.overflows && guard++ < 200) { last = addBodyPage(last); doc.recompose(); }
}
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

function rule(page, y, x1, x2) {
    var l = page.graphicLines.add();
    l.geometricBounds = [y, x1, y, x2]; l.strokeWeight = 1; l.strokeColor = BLACK;
    return l;
}

// Shrink the body-text frame on page pi to the space figures leave free.
// (Text wrap is avoided: InDesign drops some placed AI/PDF art under a wrap on export.)
function reserve(pi) {
    var tb = textBox(doc.pages[pi]), sl = slot(pi), fr = null, tc = story.textContainers;
    for (var i = 0; i < tc.length; i++)
        if (tc[i].parentPage && tc[i].parentPage.documentOffset == pi) fr = tc[i];
    if (!fr) return;
    var y1 = tb[0] + sl.top, y2 = tb[2] - sl.bottom;
    if (y2 - y1 < 3 * 12) fr.remove();          // thread skips straight to the next page
    else fr.geometricBounds = [y1, tb[1], y2, tb[3]];
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

// Number of legend paragraphs that are fully set in frame f.
function fittingParagraphs(f) {
    var ps = f.parentStory.paragraphs, n = 0;
    for (var i = 0; i < ps.length; i++) {
        var ls = ps[i].lines;
        if (ls.length && ls[ls.length - 1].parentTextFrames.length) n = i + 1; else break;
    }
    return n;
}

// Remove paragraphs [from, to) of a frame's story, leaving no stray return.
function dropParagraphs(f, from, to) {
    var st = f.parentStory, ps = st.paragraphs;
    if (to <= from) return;
    ps.itemByRange(from, to - 1).remove();
    var c = st.characters;
    if (c.length && c[-1].contents == "\r") c[-1].remove();
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
    rect.strokeWeight = 0;
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
    var s = Math.min(1, W / natW);

    // Decide how big it is and how much legend fits with it.
    var extras = CFG.figGap + CFG.ruleGap + CFG.wrapGap;
    var minLegend = Math.min(legendH, CFG.minLegendLines * lineH);
    var room = avail - slot(pi).top - slot(pi).bottom - extras;
    if (room < natH * s * 0.6 + minLegend && (slot(pi).top || slot(pi).bottom)) {
        rect.remove(); return placeFigure(fig, pi + 1);
    }
    if (natH * s + legendH > room)      // shrink a little to keep the whole legend
        s *= Math.min(1, Math.max(CFG.minFigScale, (room - legendH) / (natH * s)));
    if (natH * s + minLegend > room)    // shrink further to keep some legend with it
        s *= (room - minLegend) / (natH * s);
    var figW = natW * s, figH = natH * s;

    // Set the legend; panels that do not fit continue on the next page.
    var lf = legendFrame(page, tb, room - figH, legendFile);
    var split = lf.overflows;
    if (split) {
        var keep = Math.max(1, fittingParagraphs(lf));
        continueLegend(fig, keep, pi + 1);
        dropParagraphs(lf, keep, lf.parentStory.paragraphs.length);
        doc.recompose();
    }
    var legendHere = usedHeight(lf);

    var atTop = CFG.figurePosition == "top" ? !slot(pi).top : !!slot(pi).bottom;
    var blockH = figH + CFG.figGap + legendHere;
    var y0 = atTop ? tb[0] + slot(pi).top : tb[2] - slot(pi).bottom - blockH;
    var x0 = tb[1] + (W - figW) / 2;

    rect.geometricBounds = [y0, x0, y0 + figH, x0 + figW];
    if (file) {   // scale the whole image, then shift so the content box fills the frame
        var gw = fullW * s, gh = fullH * s;
        var gy = y0 - crop[1] * gh, gx = x0 - crop[0] * gw;
        rect.allGraphics[0].geometricBounds = [gy, gx, gy + gh, gx + gw];
    }
    lf.geometricBounds = [y0 + figH + CFG.figGap, tb[1], y0 + blockH, tb[3]];
    if (file) PLACED.push({ page: page, file: file, frame: rect.geometricBounds,
                            graphic: rect.allGraphics[0].geometricBounds, rect: rect });

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
    LOG.push("Figure " + fig.num + ": cited p" + (citePage + 1) + ", placed p" + (pi + 1) +
             (atTop ? " top" : " bottom") + ", " + Math.round(100 * s) + "% of file size" +
             (split ? ", legend continued on p" + (pi + 2) : ""));
}

// Legend paragraphs from index `from` on continue at the foot of page pi.
function continueLegend(fig, from, pi) {
    while (pi >= doc.pages.length) addBodyPage(story.textContainers[story.textContainers.length - 1]);
    var page = doc.pages[pi], tb = textBox(page);
    var cf = legendFrame(page, tb, tb[2] - tb[0], File(BUILD + "/" + fig.legend));
    var st = cf.parentStory;
    st.paragraphs.itemByRange(0, from - 1).remove();
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
for (var fi = 0; fi < M.figures.length; fi++) placeFigure(M.figures[fi]);

// Re-place every figure into a fresh frame with the final geometry. Frames that went
// through the sizing steps above can be silently dropped from the PDF export for some
// Illustrator files (an InDesign quirk); fresh frames export fine.
for (var pf = 0; pf < PLACED.length; pf++) {
    var P = PLACED[pf];
    P.rect.remove();
    var nr = P.page.rectangles.add({ geometricBounds: P.frame });
    nr.strokeWeight = 0;
    app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_PDF;
    nr.place(P.file);
    nr.allGraphics[0].geometricBounds = P.graphic;
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
doc.save(File(OUTPDF.fsName.replace(/\.pdf$/i, "") + ".indd"));
LOG.push("pages: " + doc.pages.length);
doc.close(SaveOptions.NO);
LOG.join("\n");
