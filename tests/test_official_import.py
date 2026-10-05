import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scau-thesis-format-cn" / "scripts"))
import import_official_2024_assets as importer


class OfficialImportTests(unittest.TestCase):
    def test_whole_set_verified_before_copying(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / "source"
            target = root / "official"
            source.mkdir()
            target.mkdir()
            first = source / "first.doc"
            first.write_bytes(b"verified first file")
            (target / "first.doc").write_bytes(b"existing package")
            manifest = {"required_files": [
                {"filename": "first.doc", "sha256": hashlib.sha256(first.read_bytes()).hexdigest().upper()},
                {"filename": "missing.pdf", "sha256": "0" * 64},
            ]}
            with patch.object(importer, "load_manifest", return_value=manifest), \
                    patch.object(importer, "OFFICIAL_DIR", target), \
                    patch.object(importer, "TEMPLATE_DIR", root / "template"):
                with self.assertRaises(FileNotFoundError):
                    importer.import_files(source, False)
                self.assertEqual((target / "first.doc").read_bytes(), b"existing package")
                self.assertFalse((root / "template").exists())

    def test_verify_only_leaves_source_and_target_unmodified(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / "sample.doc"
            source.write_bytes(b"fixture")
            manifest = {"required_files": [{"filename": source.name,
                "sha256": hashlib.sha256(source.read_bytes()).hexdigest().upper()}]}
            with patch.object(importer, "load_manifest", return_value=manifest), \
                    patch.object(importer, "OFFICIAL_DIR", root / "target"):
                result = importer.verify_files(root)
                self.assertEqual(len(result), 1)
                self.assertFalse((root / "target").exists())
                self.assertEqual(source.read_bytes(), b"fixture")

    def test_hash_bypass_cannot_relabel_changed_files_as_2024(self):
        with self.assertRaisesRegex(ValueError, "2024"):
            importer.verify_files(Path("."), True)


if __name__ == "__main__":
    unittest.main()
