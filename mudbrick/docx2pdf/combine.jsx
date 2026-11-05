// combine.jsx — append the supplement's pages to the preprint, as one InDesign document
// and one PDF (for uploading a single file, e.g. to bioRxiv).
//
// Run by `mudbrick docx2pdf` after layout.jsx has made both documents.
//
// arguments: [preprintIndd, supplementIndd, outputPdf]
#target indesign

app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
var LOG = [];

var OUTPDF = File(arguments[2]);
var doc = app.open(File(arguments[0]), false), supp = app.open(File(arguments[1]), false);
try {
    // One spread at a time, each after the last: copied together they arrive in reverse.
    // Each frame keeps the text it showed; the pages use the preprint's master (footer).
    var first = doc.pages.length;
    for (var i = 0; i < supp.spreads.length; i++)
        supp.spreads[i].duplicate(LocationOptions.AFTER, doc.spreads[-1]);

    // Supplement pages are numbered S1, S2, ... (the copied spreads may bring their own
    // section, which then just gets these settings).
    var p0 = doc.pages[first], sec = p0.appliedSection;
    var props = { continueNumbering: false, pageNumberStart: 1, sectionPrefix: "S", includeSectionPrefix: true };
    if (sec.pageStart == p0) sec.properties = props;
    else doc.sections.add(p0, props);

    // Export with the preset layout.jsx made, on the same ladder of image settings:
    // InDesign cannot resample some images embedded in Illustrator files.
    var preset = app.pdfExportPresets.itemByName("Preprint - Maximum Quality");
    var ladder = [[450, 600], [600, 900], null];
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
            break;
        } catch (e) {
            LOG.push("Export with " + (L ? L[0] + "/" + L[1] + " ppi" : "no downsampling") + " failed (" + e.message + "); trying gentler settings.");
            if (!L) throw e;
        }
    }
    doc.save(File(OUTPDF.fsName.replace(/\.pdf$/i, "") + ".indd"));
    LOG.push("combined: " + first + " preprint + " + (doc.pages.length - first) + " supplement pages");
} finally {
    supp.close(SaveOptions.NO);
    doc.close(SaveOptions.NO);
}
LOG.join("\n");
