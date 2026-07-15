import hashlib
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from aliyun_photo_manager.config import OssConfig
from aliyun_photo_manager.downloader import _resolve_worker_count, build_local_relative_path
from aliyun_photo_manager.update_manager import UpdateError, UpdatePackage, download_update_package


class DownloadPathSafetyTests(unittest.TestCase):
    def test_keeps_safe_relative_structure(self) -> None:
        self.assertEqual(
            build_local_relative_path("photos/2026/user/a.jpg", "photos/2026"),
            Path("user/a.jpg"),
        )

    def test_rejects_parent_traversal(self) -> None:
        with self.assertRaises(ValueError):
            build_local_relative_path("photos/../../outside.txt", "photos")

    def test_rejects_windows_style_and_drive_paths(self) -> None:
        for key in (r"photos\..\outside.txt", "C:/outside.txt"):
            with self.subTest(key=key), self.assertRaises(ValueError):
                build_local_relative_path(key, "")

    def test_streaming_download_honors_requested_worker_count(self) -> None:
        config = OssConfig("aliyun", "id", "secret", "endpoint", "bucket")
        self.assertEqual(_resolve_worker_count(config, None, 12), 12)


class UpdateIntegrityTests(unittest.TestCase):
    def _package(self, source: Path, sha256: str) -> UpdatePackage:
        return UpdatePackage(
            version="9.9.9",
            package_type="full",
            url=source.as_uri(),
            notes="",
            sha256=sha256,
        )

    def test_accepts_matching_hash(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "update.zip"
            content = b"verified update"
            source.write_bytes(content)
            result = download_update_package(self._package(source, hashlib.sha256(content).hexdigest()))
            self.assertEqual(result.read_bytes(), content)

    def test_deletes_package_when_hash_mismatches(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "update.zip"
            source.write_bytes(b"tampered update")
            with self.assertRaises(UpdateError):
                download_update_package(self._package(source, "0" * 64))


if __name__ == "__main__":
    unittest.main()
