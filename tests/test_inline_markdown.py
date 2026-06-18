"""Tests for cv_renderer.inline_markdown.

Run with:  python -m unittest tests.test_inline_markdown

No pytest dependency — stdlib unittest only.
"""
import sys
import unittest
from pathlib import Path

# Make project root importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from cv_renderer.inline_markdown import to_html, to_plain, iter_docx_runs, tokenize


class TestToHtml(unittest.TestCase):
    def test_plain_passes_through_escaped(self):
        self.assertEqual(to_html("plain text"), "plain text")

    def test_bold_becomes_strong(self):
        self.assertEqual(to_html("**bold**"), "<strong>bold</strong>")

    def test_italic_becomes_em(self):
        self.assertEqual(to_html("*italic*"), "<em>italic</em>")

    def test_link_becomes_anchor(self):
        self.assertEqual(
            to_html("[label](https://x.com)"),
            '<a href="https://x.com">label</a>',
        )

    def test_html_chars_are_escaped(self):
        self.assertEqual(
            to_html("a < b & c > d"),
            "a &lt; b &amp; c &gt; d",
        )

    def test_bold_inside_sentence(self):
        self.assertEqual(
            to_html("- **Bachelor's degree** in Communication"),
            "- <strong>Bachelor&#x27;s degree</strong> in Communication",
        )

    def test_multiple_marks_mixed(self):
        out = to_html("**A** then *B* then [C](http://c)")
        self.assertEqual(
            out,
            '<strong>A</strong> then <em>B</em> then <a href="http://c">C</a>',
        )

    def test_empty_input_safe(self):
        self.assertEqual(to_html(""), "")
        self.assertEqual(to_html(None), "")  # type: ignore[arg-type]


class TestToPlain(unittest.TestCase):
    def test_bold_strips_asterisks(self):
        self.assertEqual(to_plain("**bold**"), "bold")

    def test_italic_strips_asterisks(self):
        self.assertEqual(to_plain("*italic*"), "italic")

    def test_link_keeps_url(self):
        self.assertEqual(
            to_plain("[label](https://x.com)"), "label (https://x.com)"
        )

    def test_link_with_same_url_as_label_no_duplicate(self):
        self.assertEqual(
            to_plain("[https://x.com](https://x.com)"), "https://x.com"
        )

    def test_html_chars_kept_literal(self):
        self.assertEqual(to_plain("a < b & c"), "a < b & c")


class TestDocxRuns(unittest.TestCase):
    def test_yields_typed_tuples(self):
        runs = list(iter_docx_runs("Hello **world**"))
        self.assertEqual(
            runs,
            [("Hello ", False, False), ("world", True, False)],
        )

    def test_italic_run(self):
        runs = list(iter_docx_runs("a *b* c"))
        self.assertEqual(
            runs,
            [("a ", False, False), ("b", False, True), (" c", False, False)],
        )

    def test_link_flattens_to_text_url(self):
        runs = list(iter_docx_runs("see [docs](https://d)"))
        # Note: link runs are emitted as plain (False, False) for now;
        # python-docx hyperlinks would require OOXML plumbing.
        self.assertIn(("docs (https://d)", False, False), runs)


class TestTokenizerEdges(unittest.TestCase):
    def test_unbalanced_asterisks_kept_as_text(self):
        # Single trailing * is not a valid italic open/close — stays literal.
        toks = tokenize("hello *world")
        kinds = [t[0] for t in toks]
        self.assertNotIn("italic", kinds)
        self.assertNotIn("bold", kinds)

    def test_nested_not_supported_but_doesnt_crash(self):
        # Bold inside link is intentionally NOT parsed; should not raise.
        out = to_html("[**bold-label**](http://x)")
        self.assertIn("<a", out)


if __name__ == "__main__":
    unittest.main()
