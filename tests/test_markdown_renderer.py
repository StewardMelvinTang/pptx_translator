import unittest

import pptxtranslator as app


class RecordingText:
    def __init__(self):
        self.fragments = []
        self.configured_tags = {}
        self.bindings = []

    def insert(self, _position, text, tags=()):
        if isinstance(tags, str):
            tags = (tags,)
        self.fragments.append((text, tuple(tags)))

    def tag_configure(self, name, **options):
        self.configured_tags[name] = options

    def tag_bind(self, name, event, callback):
        self.bindings.append((name, event, callback))


class MarkdownRendererTests(unittest.TestCase):
    def test_renders_common_markdown_without_control_markers(self):
        chat = app.ChatPanel.__new__(app.ChatPanel)
        surface = RecordingText()
        markdown = (
            "### Summary\n"
            "Use **bold**, *italic*, `code`, and [OpenAI](https://openai.com).\n\n"
            "- First item\n"
            "> A note\n\n"
            "```python\nprint('hello')\n```\n"
        )

        chat._insert_markdown(surface, markdown)

        rendered = "".join(text for text, _tags in surface.fragments)
        all_tags = {tag for _text, tags in surface.fragments for tag in tags}
        self.assertIn("Summary", rendered)
        self.assertIn("print('hello')", rendered)
        self.assertNotIn("```", rendered)
        self.assertTrue({"h3", "bold", "italic", "inline_code", "code"} <= all_tags)
        self.assertTrue(any(tag.startswith("link_") for tag in all_tags))


if __name__ == "__main__":
    unittest.main()
