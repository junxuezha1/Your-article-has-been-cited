import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from tools.installer_build import (
    PROJECT_FILES,
    find_iscc,
    make_inno_script,
    make_installed_launcher,
    should_skip_source_path,
)


class InstallerBuildTests(unittest.TestCase):
    def test_source_excludes_local_data_and_archive_materials(self):
        skipped = [
            Path("data/input/example.docx"),
            Path("数据/persons.xls"),
            Path("整理归档/发布压缩包/archive.zip"),
            Path("软件著作权申请资料/截图/1.png"),
            Path("邮箱检索/out/result.csv"),
            Path("__pycache__/app.pyc"),
            Path(".git/config"),
        ]
        included = [
            Path("app.py"),
            Path("templates/index.html"),
            Path("config.example.yaml"),
            Path("启动系统.bat"),
            Path("requirements.txt"),
        ]

        for path in skipped:
            self.assertTrue(should_skip_source_path(path), path)
        for path in included:
            self.assertFalse(should_skip_source_path(path), path)

    def test_project_manifest_contains_runtime_entrypoints(self):
        self.assertIn("app.py", PROJECT_FILES)
        self.assertIn("main.py", PROJECT_FILES)
        self.assertIn("requirements.txt", PROJECT_FILES)
        self.assertIn("config.example.yaml", PROJECT_FILES)
        self.assertIn("启动系统.bat", PROJECT_FILES)
        self.assertIn("templates", PROJECT_FILES)
        self.assertIn("assets", PROJECT_FILES)

    def test_installed_launcher_uses_private_python_and_user_data(self):
        launcher = make_installed_launcher()

        self.assertIn(r"runtime\python\python.exe", launcher)
        self.assertIn("CITATION_NOTIFIER_HOME", launcher)
        self.assertIn(r"%LOCALAPPDATA%\CitationNotifier", launcher)
        self.assertIn("main.py web", launcher)

    def test_find_iscc_checks_path_and_common_install_locations(self):
        with tempfile.TemporaryDirectory() as tmp:
            fake = Path(tmp) / "ISCC.exe"
            fake.write_text("", encoding="utf-8")
            with patch("tools.installer_build.shutil.which", return_value=None):
                self.assertEqual(find_iscc([fake.parent]), fake)

    def test_inno_script_does_not_require_optional_chinese_language_pack(self):
        script = make_inno_script()

        self.assertIn("[Setup]", script)
        self.assertNotIn("ChineseSimplified.isl", script)


if __name__ == "__main__":
    unittest.main()
