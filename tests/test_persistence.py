import tempfile
import unittest
from pathlib import Path

import pptxtranslator as app


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.old_history_path = app.HISTORY_PATH
        self.old_chat_dir = app.CHAT_SESSIONS_DIR
        app.HISTORY_PATH = self.root / "history.json"
        app.CHAT_SESSIONS_DIR = self.root / "chats"

    def tearDown(self):
        app.HISTORY_PATH = self.old_history_path
        app.CHAT_SESSIONS_DIR = self.old_chat_dir
        self.temp_dir.cleanup()

    def test_translation_history_round_trip(self):
        records = [{"id": "record-1", "output_dir": str(self.root)}]
        self.assertTrue(app.save_translation_history(records))
        self.assertEqual(app.load_translation_history(), records)

    def test_chat_session_is_scoped_to_file_and_excludes_system_context(self):
        first = self.root / "first.pptx"
        second = self.root / "second.pptx"
        first.touch()
        second.touch()
        messages = [
            {"role": "system", "content": "large presentation context"},
            {"role": "user", "content": "Summarize this file"},
            {"role": "assistant", "content": "Here is the summary"},
        ]

        self.assertTrue(app.save_chat_session(first, messages))
        self.assertEqual(len(app.load_chat_session(first)), 2)
        self.assertEqual(app.load_chat_session(second), [])


if __name__ == "__main__":
    unittest.main()
