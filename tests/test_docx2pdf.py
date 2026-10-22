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


def write_docx(path, paragraphs):
    with zipfile.ZipFile(path, "w") as z:
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
    para(run("DISCUSSION", b=True)),
    para(run("Discussion text.")),
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

    def test_figure_choice(self):
        fig = self.m["figures"][0]
        self.assertEqual(os.path.basename(fig["file"]), "fig_v2.png")   # highest version, no panels/tmp

    def test_legend_from_captions_doc(self):
        legend = self.read("legend_1.txt")
        self.assertIn("<pstyle:LegendTitle><cstyle:cs_sb>Figure 1. A figure title.", legend)
        self.assertEqual(legend.count("<pstyle:Legend>"), 2)   # (A) and the sentence-initial (B)
        self.assertNotIn("TODO", legend)                         # notes between figures dropped
        self.assertTrue(os.path.exists(os.path.join(self.dir, "paper_figure_legends.docx")))

    def test_manifest(self):
        with open(os.path.join(self.build, "manifest.json")) as fh:
            m = json.load(fh)
        self.assertEqual([f["num"] for f in m["figures"]], [1])


if __name__ == "__main__":
    unittest.main()
