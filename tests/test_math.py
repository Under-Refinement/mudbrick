"""Tests for Word-equation handling in docx2pdf (no InDesign, pandoc or LaTeX needed).

Run:  python3 -m unittest discover tests
"""
import os
import tempfile
import unittest
import zipfile

from mudbrick.docx2pdf import convert, omml, render_math

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def mr(text, sty=None):
    pr = '<m:rPr><m:sty m:val="%s"/></m:rPr>' % sty if sty else ""
    return "<m:r>%s<m:t>%s</m:t></m:r>" % (pr, text)


def sub(base, script):
    return "<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>" % (base, script)


def frac(num, den):
    return "<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>" % (num, den)


def math(*parts):
    return "<m:oMath>%s</m:oMath>" % "".join(parts)


def text(t):
    return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % t


def docx(path, paragraphs):
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
                   '<?xml version="1.0"?><w:document xmlns:w="%s" xmlns:m="%s"><w:body>%s</w:body></w:document>'
                   % (W_NS, M_NS, "".join("<w:p>%s</w:p>" % p for p in paragraphs)))


class MathParsing(unittest.TestCase):
    def read(self, *pieces):
        """One paragraph made of the given pieces (runs and equations)."""
        d = tempfile.mkdtemp()
        path = os.path.join(d, "m.docx")
        docx(path, ["".join(pieces)])
        convert.MATH.clear()
        return convert.read_paragraphs(path)

    def test_simple_inline_becomes_text(self):
        p = self.read(text("approximation "), math(sub(mr("z"), mr("i")), mr("=0")), text(" to"))[0]
        self.assertNotIn("@@MATH", convert.ptext(p))
        self.assertIn("z", convert.ptext(p))
        runs = [(t, f) for t, f in p["runs"] if t.strip()]
        self.assertTrue(any(t == "z" and f["i"] for t, f in runs))          # italic variable
        self.assertTrue(any(t == "i" and f["sub"] for t, f in runs))        # real subscript
        self.assertIn("\u2009=\u2009", convert.ptext(p))                    # spaced relation

    def test_complex_inline_becomes_placeholder(self):
        p = self.read(text("ratio "), math(frac(mr("a"), mr("b"))), text(" here"))[0]
        self.assertRegex(convert.ptext(p), omml.TOKEN_RE)
        self.assertFalse(p.get("eq"))

    def test_display_equation_with_number(self):
        p = self.read(text("\t"), math(mr("y"), mr("="), mr("x")), text("\t(3)"))[0]
        self.assertTrue(p.get("eq"))
        self.assertEqual(p["eqnum"], "(3)")
        self.assertTrue(convert.MATH[p["eq"]]["display"])

    def test_upright_and_bold_styles(self):
        runs = omml.to_runs(omml_el(math(mr("diag", "p"), mr("u", "b"))))
        self.assertEqual([(t, f["i"], f["b"]) for t, f in runs], [("diag", False, False), ("u", False, True)])

    def test_script_letters(self):
        def scr(text, kind):
            return '<m:r><m:rPr><m:scr m:val="%s"/><m:sty m:val="p"/></m:rPr><m:t>%s</m:t></m:r>' % (kind, text)
        self.assertEqual(omml.to_runs(omml_el(math(scr("l", "script"))))[0][0], "\u2113")   # script l -> text
        self.assertEqual(omml.to_runs(omml_el(math(scr("R", "double-struck"))))[0][0], "\u211d")
        self.assertIsNone(omml.to_runs(omml_el(math(sub(scr("L", "script"), mr("reproj"))))))  # -> LaTeX

    def test_key_is_stable(self):
        a = omml_el(math(mr("x"), mr("+1")))
        b = omml_el(math(mr("x"), mr("+1")))
        self.assertEqual(omml.key(a), omml.key(b))

    def test_rendered_pdfs_linked_from_equations_folder(self):
        d = tempfile.mkdtemp()
        os.makedirs(os.path.join(d, "equations"))
        open(os.path.join(d, "equations", "abc.pdf"), "wb").close()
        with open(os.path.join(d, "equations", "equations.json"), "w") as fh:
            fh.write('{"abc": {"file": "abc.pdf", "w": 10, "h": 5, "d": 1}}')
        convert.MATH.clear()
        convert.MATH["abc"] = {"runs": None, "display": True}
        build, manifest = tempfile.mkdtemp(), {}
        convert.attach_math(manifest, os.path.join(d, "m.docx"), build, {}, {"abc"})
        self.assertEqual(manifest["math"]["abc"]["file"], os.path.join(d, "equations", "abc.pdf"))
        self.assertFalse(os.path.exists(os.path.join(build, "math")))   # nothing copied


def omml_el(xml):
    import xml.etree.ElementTree as ET
    return ET.fromstring('<m:wrap xmlns:m="%s" xmlns:w="%s">%s</m:wrap>' % (M_NS, W_NS, xml))[0]


class LatexCleanup(unittest.TestCase):
    def test_word_numbering_array_unwrapped(self):
        self.assertEqual(render_math.clean("\\begin{array}{r}\nx = y,\n\\end{array}"), "x = y,")

    def test_alignment_spaces_and_upright_words(self):
        self.assertEqual(render_math.clean(r"{\ \ \ \ f}_{i} = {diag}(u)"), r"{ f}_{i} = {\mathrm{diag}}(u)")

    def test_labelled_brace(self):
        t = r"\underset{\text{error}}{\overset{a + b}{" + "\ufe38" + "}}"
        self.assertEqual(render_math.clean(t), r"\underbrace{a + b}_{\text{error}}")

    def test_line_break_after_control_space(self):
        for v in render_math.line_breaks(r"L = a + b\  + c"):
            self.assertNotIn(r"\\\&", v)                     # no literal "&" on line 2

    def test_line_breaks_split_pairs(self):
        v = render_math.line_breaks(r"a = \begin{bmatrix} 1 \\ 2 \end{bmatrix},\quad\quad b = c")
        self.assertTrue(any(r"\\b = c" in x for x in v))


if __name__ == "__main__":
    unittest.main()
