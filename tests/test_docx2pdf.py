"""Smoke tests for the docx2pdf converter (no InDesign needed).

Run:  python3 -m unittest discover tests
"""
import json
import os
import tempfile
import unittest
import zipfile

from mudbrick.docx2pdf import convert

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def run(text, b=False, i=False, sup=False):
    props = ("<w:b/>" if b else "") + ("<w:i/>" if i else "") + \
        ('<w:vertAlign w:val="superscript"/>' if sup else "")
    return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (props, text)


def para(*runs):
    return "<w:p>%s</w:p>" % "".join(runs)


PNG_1PX = bytes.fromhex("89504e470d0a1a0a0000000d4948445200000001000000010806000000"
                        "1f15c4890000000d49444154789c6360000002000154a24f5d0000000049454e44ae426082")


def picture(rid, deleted=False):
    r = ('<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
         '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>'
         '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>'
         '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="%s"/>'
         '</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' % rid)
    return '<w:del w:id="1" w:author="x">%s</w:del>' % r if deleted else r


def table(rows):
    """rows: [[(text, span, bold)]]"""
    out = "<w:tbl><w:tblGrid>%s</w:tblGrid>" % ('<w:gridCol w:w="1000"/>' * max(
        sum(c[1] for c in r) for r in rows))
    for r in rows:
        out += "<w:tr>"
        for text, span, bold in r:
            pr = '<w:tcPr><w:gridSpan w:val="%d"/></w:tcPr>' % span if span > 1 else ""
            out += "<w:tc>%s%s</w:tc>" % (pr, para(run(text, b=bold)))
        out += "</w:tr>"
    return out + "</w:tbl>"


def write_docx(path, paragraphs, media=()):
    with zipfile.ZipFile(path, "w") as z:
        for name in media:
            z.writestr("word/media/" + name, PNG_1PX)
        z.writestr("word/_rels/document.xml.rels",
                   '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
                   % "".join('<Relationship Id="rId%d" Type="image" Target="media/%s"/>' % (i + 1, n)
                             for i, n in enumerate(media)))
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml",
                   '<?xml version="1.0"?><w:document %s><w:body>%s</w:body></w:document>'
                   % (NS, "".join(paragraphs)))


MANUSCRIPT = [
    para(run("[TITLE]", b=True)),
    para(run("Ada Lovelace"), run("1", sup=True)),
    para(run("1", sup=True), run("Department of Biology,")),
    para(run("KEYWORDS:", b=True)),
    para(run("RUNNING TITLE:", b=True)),
    para(run("A short running title")),
    para(run("ABSTRACT", b=True)),
    para(run("An abstract.")),
    para(run("RESULTS", b=True)),
    para(run("A subheading.", b=True)),
    para(run("Results text citing (Figure 1A) and (Supplementary Figure 3), "
             "as shown by Smith et al. 2020.")),
    para(run("[editor stopped here]")),
    para(run("L = λ"), '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>dice</w:t></w:r>', run(" L")),
    para(run("with λ = 2.0, which was chosen after a careful sweep of values.")),
    para(run("DISCUSSION", b=True)),
    para(run("Discussion text.")),
    para(run("data availability ", b=True), run("Data are on Zenodo.")),
    para(run("REFERENCES", b=True)),
    para(run("Smith J. 2020. A paper. "), run("Journal", i=True), run(" "), run("1", b=True), run(": 1-2.")),
]

CAPTIONS = [
    para(run("Figure 1: A figure title.", b=True), run(" "),
         run("(A)", b=True), run(" Panel A, compare with "), run("(B)", b=True),
         run(" here. "), run("(B)", b=True), run(" Panel B.")),
    para(run("TODO: replace panel B")),
]


class Docx2PdfConvert(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.docx = os.path.join(self.dir, "paper.docx")
        write_docx(self.docx, MANUSCRIPT)
        write_docx(os.path.join(self.dir, "paper_figures.docx"), CAPTIONS)
        os.makedirs(os.path.join(self.dir, "figs", "figure_1"))
        for name in ("fig_v1.png", "fig_v2.png", "fig_v2_panel_A.png", "~ai-123.tmp"):
            open(os.path.join(self.dir, "figs", "figure_1", name), "wb").close()
        self.build = os.path.join(self.dir, "build")
        cfg = dict(convert.DEFAULT_CONFIG, trim_figures=False)
        self.m = convert.convert(self.docx, self.build, cfg, os.path.join(self.dir, "figs"))

    def read(self, name):
        with open(os.path.join(self.build, name)) as fh:
            return fh.read()

    def test_front_matter(self):
        self.assertEqual(self.m["running_title"], "A short running title")
        self.assertIn("<pstyle:Title>", self.read("title.txt"))
        self.assertTrue(any("placeholder" in w for w in self.m["warnings"]))
        self.assertNotIn("Keywords", self.read("abstract.txt"))   # empty keywords omitted

    def test_body(self):
        body = self.read("body.txt")
        self.assertIn("<pstyle:H2>", body)
        self.assertIn("<cstyle:cs_i>et al.<cstyle:>", body)
        self.assertTrue(any("editor stopped here" in w for w in self.m["warnings"]))
        self.assertLess(body.index("REFERENCES"), len(body))

    def test_run_in_back_matter(self):
        body = self.read("body.txt")
        self.assertIn("<pstyle:H1Back>DATA AVAILABILITY", body)
        self.assertLess(body.index("REFERENCES"), body.index("DATA AVAILABILITY"))
        self.assertIn("<pstyle:Body>Data are on Zenodo.", body)

    def test_panels_on_own_lines_when_configured(self):
        cfg = dict(convert.DEFAULT_CONFIG, trim_figures=False, legend_panels_inline=False)
        convert.convert(self.docx, self.build, cfg, os.path.join(self.dir, "figs"))
        self.assertEqual(self.read("legend_1.txt").count("<pstyle:Legend>"), 2)

    def test_uniform_spacing(self):
        self.assertNotIn("BodyTight", self.read("body.txt"))

    def test_equations_centered(self):
        body = self.read("body.txt")
        self.assertIn("<pstyle:Equation>L = ", body)
        self.assertNotIn("<pstyle:Equation>with", body)

    def test_word_equation_object(self):
        p = convert._read_para(convert.ET.fromstring(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathPara><m:oMath>'
            '<m:r><m:t>E=m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e>'
            '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara></w:p>'),
            {}, "accept", {})
        self.assertTrue(p["math"])
        self.assertEqual(convert.ptext(p), "E=mc2")
        self.assertTrue(p["runs"][-1][1]["sup"])

    def test_caption_line_break_and_position_labels(self):
        legend = convert.normalize_legend([[("Figure 2. A title.", {"b": True}), ("\n", {}),
                                            ("(left)", {"b": True}), (" Left plot. ", {}),
                                            ("(right)", {"b": True}), (" Right plot.", {})]])
        self.assertEqual(legend[0][0], "LegendTitle")
        self.assertEqual("".join(t for t, _ in legend[0][1]).strip(), "Figure 2. A title.")
        self.assertNotIn("\n", "".join(t for _, r in legend for t, _ in r))

    def test_figure_choice(self):
        fig = self.m["figures"][0]
        self.assertEqual(os.path.basename(fig["file"]), "fig_v2.png")   # highest version, no panels/tmp

    def test_legend_from_captions_doc(self):
        legend = self.read("legend_1.txt")
        self.assertIn("<pstyle:LegendTitle><cstyle:cs_sb>Figure 1. A figure title.", legend)
        self.assertEqual(legend.count("<pstyle:Legend>"), 1)   # panels run on after the title
        self.assertIn("here. <cstyle:cs_sb>(B)<cstyle:> Panel B.", legend)
        self.assertNotIn("TODO", legend)                         # notes between figures dropped
        self.assertTrue(os.path.exists(os.path.join(self.dir, "paper_figure_legends.docx")))

    def test_manifest(self):
        with open(os.path.join(self.build, "manifest.json")) as fh:
            m = json.load(fh)
        self.assertEqual([f["num"] for f in m["figures"]], [1])


if __name__ == "__main__":
    unittest.main()


class Docx2PdfThreeDocuments(unittest.TestCase):
    """Manuscript + captions .docx with embedded pictures + supplementary .docx."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.docx = os.path.join(self.dir, "paper.docx")
        write_docx(self.docx, MANUSCRIPT)
        write_docx(os.path.join(self.dir, "paper_maintext_figures.docx"), [
            para(picture("rId1", deleted=True), picture("rId2")),   # old picture tracked-deleted
            para(run("Figure 1. A title.", b=True), run(" Text.")),
        ], media=("old.png", "new.png"))
        write_docx(os.path.join(self.dir, "paper_supp_materials.docx"), [
            para(picture("rId1")),
            para(run("Supplementary Figure. 1.", b=True), run(" A supplementary figure.")),
            para(run("Table 1: A table.", b=True)),
            table([[("Overall", 2, True)], [("A", 1, True), ("B", 1, True)], [("1", 1, False), ("2", 1, False)]]),
        ], media=("s1.png",))
        self.build = os.path.join(self.dir, "build")

    def test_supplementary_figure_files(self):
        figs = os.path.join(self.dir, "figs")
        for rel in ("supp_figure_1/s1_v1.ai", "supp_figure_1/s1_v2.ai", "figure_S2/a.ai",
                    "supplementary/figure_3/b.ai", "figure_1/main.ai"):
            os.makedirs(os.path.dirname(os.path.join(figs, rel)), exist_ok=True)
            open(os.path.join(figs, rel), "wb").close()
        open(os.path.join(figs, "Supp_Fig4.ai"), "wb").close()
        name = lambda n: os.path.basename(convert.find_figure(figs, n, supplementary=True) or "")
        self.assertEqual([name(1), name(2), name(3), name(4)],
                         ["s1_v2.ai", "a.ai", "b.ai", "Supp_Fig4.ai"])
        self.assertEqual(name(9), "")
        # main-figure search finds main figures only, never the supplement's
        self.assertEqual(os.path.basename(convert.find_figure(figs, 1)), "main.ai")
        self.assertIsNone(convert.find_figure(figs, 3))
        self.assertIsNone(convert.find_figure(figs, 4))

    def test_figure_folder_contents_never_match_other_figures(self):
        figs = os.path.join(self.dir, "figs")
        for rel in ("supp_figure_1/figure4_old.ai", "figure_2/Fig3_maps.ai"):
            os.makedirs(os.path.dirname(os.path.join(figs, rel)))
            open(os.path.join(figs, rel), "wb").close()
        self.assertIsNone(convert.find_figure(figs, 4, supplementary=True))
        self.assertIsNone(convert.find_figure(figs, 3))
        self.assertEqual(os.path.basename(convert.find_figure(figs, 2)), "Fig3_maps.ai")

    def test_supplement_prefers_file_over_pasted_picture(self):
        figs = os.path.join(self.dir, "figs")
        os.makedirs(figs)
        with open(os.path.join(figs, "FigS1.png"), "wb") as fh:
            fh.write(PNG_1PX)
        supp = convert.find_supplement_docx(self.docx)
        cfg = dict(convert.DEFAULT_CONFIG, trim_figures=False)
        fig = convert.convert_supplement(supp, self.docx, self.build, cfg, "RT", figs)["si_figures"][0]
        self.assertEqual((fig["source"], os.path.basename(fig["file"])), ("file", "FigS1.png"))
        fig = convert.convert_supplement(supp, self.docx, self.build,
                                         dict(cfg, figures_from_captions=True), "RT", figs)["si_figures"][0]
        self.assertEqual(fig["source"], "pasted")

    def test_supplement_only_dry_run(self):
        from mudbrick.cli import main
        self.assertEqual(main(["docx2pdf", self.docx, "--supplement-only", "--dry-run"]), 0)
        self.assertFalse(os.path.exists(os.path.join(self.dir, "paper_preprint.pdf")))
        self.assertEqual(main(["docx2pdf", self.docx, "--supplement-only", "--no-supplement"]), 2)

    def test_figures_from_captions_document(self):
        m = convert.convert(self.docx, self.build, dict(convert.DEFAULT_CONFIG, trim_figures=False),
                            self.dir, write_outputs=False)
        fig = m["figures"][0]
        self.assertEqual(fig["source"], "captions document")
        self.assertTrue(fig["file"].endswith("embedded_fig1.png"))
        self.assertFalse(os.path.exists(os.path.join(self.dir, "paper_figure_legends.docx")))  # dry run

    def test_supplement(self):
        supp = convert.find_supplement_docx(self.docx)
        self.assertTrue(supp.endswith("paper_supp_materials.docx"))
        self.assertTrue(convert.find_legends_docx(self.docx, convert.DEFAULT_CONFIG)
                        .endswith("paper_maintext_figures.docx"))
        m = convert.convert_supplement(supp, self.docx, self.build, dict(convert.DEFAULT_CONFIG), "RT")
        self.assertEqual(len(m["si_figures"]), 1)
        self.assertEqual(m["tables"][0]["rows"][0][0]["span"], 2)
        with open(os.path.join(self.build, "si.txt")) as fh:
            si = fh.read()
        # figures get their own pages; each legend is a separate frame, not in the story
        self.assertIn("[[FIGPAGES]]", si)
        self.assertNotIn("Supplementary Figure 1.", si)
        self.assertEqual(m["si_figures"][0]["legend"], "si_legend_1.txt")
        with open(os.path.join(self.build, "si_legend_1.txt")) as fh:
            legend = fh.read()
        with open(os.path.join(self.build, "si_figures_head.txt")) as fh:
            self.assertIn("SUPPLEMENTARY FIGURES", fh.read())
        self.assertIn("Supplementary Table 1.", si)
        self.assertNotIn("<pstyle:LegendTitle>", legend)   # title runs into the caption text
        with open(os.path.join(self.build, "cover.txt")) as fh:
            cover = fh.read()
        self.assertIn("<pstyle:SICoverHead>Supplementary Information", cover)
        self.assertIn("<pstyle:AffilFirst>1. Department of Biology", cover)  # same as main title page
        self.assertLess(si.index("[[TABLE:0]]"), si.index("Supplementary Table 1."))  # caption below
        self.assertLess(si.index("[[FIGPAGES]]"), si.index("SUPPLEMENTARY TABLES"))  # tables last
        self.assertIn("<pstyle:SICaption><cstyle:cs_sb>Supplementary Figure 1.", legend)

    def test_supplement_tables_from_their_own_file(self):
        supp = convert.find_supplement_docx(self.docx)
        tables = os.path.join(self.dir, "figs", "Tables S1-2.docx")
        os.makedirs(os.path.dirname(tables))
        write_docx(tables, [
            para(run("Table S1. ", b=True), run("First table")),
            table([[("Origin", 1, True), ("1", 1, False)], [("x", 1, False), ("2", 1, False)]]),
            para(run("*a note under the table.")),
            para(run("Table S2. ", b=True), run("Second table")),
            table([[("A", 1, True), ("B", 1, True)], [("c", 1, False), ("d", 1, False)]])])
        # a superscript inside a cell is kept as a run of its own
        with zipfile.ZipFile(tables) as z:
            doc = z.read("word/document.xml").decode()
        doc = doc.replace('<w:t xml:space="preserve">1</w:t>',
                          '<w:t xml:space="preserve">1</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>st</w:t>', 1)
        write_raw = {}
        with zipfile.ZipFile(tables) as z:
            for n in z.namelist():
                write_raw[n] = z.read(n)
        write_raw["word/document.xml"] = doc.encode()
        with zipfile.ZipFile(tables, "w") as z:
            for n, d in write_raw.items():
                z.writestr(n, d)
        self.assertEqual(convert.find_tables_docx(self.docx, os.path.dirname(tables)), tables)
        m = convert.convert_supplement(supp, self.docx, self.build, dict(convert.DEFAULT_CONFIG),
                                       "RT", os.path.dirname(tables))
        self.assertEqual([len(t["rows"]) for t in m["tables"]], [2, 2])   # from the tables file only
        cell = m["tables"][0]["rows"][0][1]
        self.assertEqual(cell["t"], "1st")
        self.assertEqual(cell["runs"], [["1", []], ["st", ["sup"]]])
        with open(os.path.join(self.build, "si.txt")) as fh:
            si = fh.read()
        self.assertIn("a note under the table", si)
        self.assertTrue(any("ignored" in w for w in m["warnings"]))  # supplement's own table

    def test_supplement_labels_follow_manuscript_citations(self):
        cite = lambda text: [{"runs": [(text, {})]}]
        self.assertEqual(convert.supplement_label_style(cite("as in Figure S2 and Table S1")), "S")
        self.assertEqual(convert.supplement_label_style(cite("see Supplementary Figure 2")), "long")
        self.assertEqual(convert.supplement_label("Table", 3, "S"), "Table S3.")


class Equations(unittest.TestCase):
    """Word equations (OMML) -> LaTeX; no TeX engine needed."""
    MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def tex(self, inner):
        import xml.etree.ElementTree as ET
        from mudbrick.docx2pdf import equations
        el = ET.fromstring('<m:oMath xmlns:m="%s" %s>%s</m:oMath>' % (self.MNS, NS, inner))
        return equations.omml_to_tex(el)

    @staticmethod
    def r(t):
        return "<m:r><m:t>%s</m:t></m:r>" % t

    def test_subscripts_fractions_sums(self):
        r = self.r
        eq = ("<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>%s"
              "<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>"
              '<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/>'
              '<m:supHide m:val="1"/></m:naryPr><m:sub>%s</m:sub><m:sup/><m:e>%s</m:e></m:nary>'
              % (r("L"), r("focal"), r("=-"), r("1"), r("N"), r("i"), r("log")))
        self.assertEqual(self.tex(eq),
                         "{L}_{\\mathrm{focal}}=−\\frac{1}{N}\\sum\\limits_{i}{\\log }")

    def test_delimiters_and_tracked_deletions(self):
        r = self.r
        eq = ('<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
              "<m:e>%s</m:e></m:d><m:r><w:del><m:t>x</m:t></w:del></m:r>" % r("a+b"))
        self.assertEqual(self.tex(eq), "\\left[ a+b \\right]")
