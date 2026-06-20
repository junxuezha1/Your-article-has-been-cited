import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import app as app_module
import main as main_module
from extract_references import split_reference_rows
from send_emails import render_email


class MultiJournalBehaviorTests(unittest.TestCase):
    def test_reference_split_keeps_only_chinese_journals_as_papers(self):
        rows = [
            {"authors": "张三", "title": "中文期刊论文", "journal": "高等教育研究", "doc_type": "J", "raw_text": "张三. 中文期刊论文[J]. 高等教育研究,2024(1)."},
            {"authors": "李四", "title": "一本书", "journal": "北京: 高等教育出版社", "doc_type": "M", "raw_text": "李四. 一本书[M]. 北京: 高等教育出版社,2020."},
            {"authors": "教育部", "title": "教育政策文件", "journal": "", "doc_type": "Z", "raw_text": "教育部. 教育政策文件[Z]."},
            {"authors": "Smith J", "title": "Foreign article", "journal": "Journal of Education", "doc_type": "J", "raw_text": "Smith J. Foreign article[J]. Journal of Education, 2023."},
        ]

        papers, other = split_reference_rows(rows)

        self.assertEqual([row["title"] for row in papers], ["中文期刊论文"])
        self.assertEqual({row["title"] for row in other}, {"一本书", "教育政策文件", "Foreign article"})

    def test_email_content_uses_current_journal_metadata(self):
        recipient = {
            "email": "author@example.com",
            "author_name": "张三",
            "citations": [
                {
                    "cited_paper_title": "被引论文",
                    "cited_paper_authors": "张三",
                    "citing_paper_title": "来源论文",
                    "citing_paper_authors": "李四",
                    "citing_paper_file": "source.docx",
                }
            ],
        }
        journal = {
            "name": "中南大学学报（社会科学版）",
            "name_en": "Journal of Central South University (Social Science)",
            "intro": "综合性人文社科学术期刊。",
            "editorial_office": "《中南大学学报（社会科学版）》编辑部",
        }

        subject, html = render_email(recipient, "templates/notification.html", journal)

        self.assertIn("中南大学学报（社会科学版）", subject)
        self.assertIn("综合性人文社科学术期刊", html)
        self.assertIn("《中南大学学报（社会科学版）》编辑部", html)

    def test_flask_journal_switch_changes_theme_and_data_scope(self):
        old_base = app_module.BASE_DIR
        old_runtime = app_module.RUNTIME_DIR
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                app_module.BASE_DIR = tmp_path
                app_module.RUNTIME_DIR = tmp_path
                (tmp_path / "templates").mkdir()
                (tmp_path / "templates" / "notification.html").write_text(
                    (old_base / "templates" / "notification.html").read_text(encoding="utf-8"),
                    encoding="utf-8",
                )
                for journal_id, title in [("innovation", "双创被引论文"), ("csu_social", "社科被引论文")]:
                    data_dir = tmp_path / "data" / journal_id
                    data_dir.mkdir(parents=True)
                    pd.DataFrame([
                        {
                            "email": f"{journal_id}@example.com",
                            "matched_name": "张三",
                            "title": title,
                            "authors": "张三",
                            "article_title": "来源论文",
                            "article_authors": "李四",
                            "source_file": "source.docx",
                        }
                    ]).to_csv(data_dir / "authors_emails.csv", index=False, encoding="utf-8-sig")

                client = app_module.app.test_client()
                innovation_html = client.get("/send?journal=innovation").get_data(as_text=True)
                social_html = client.get("/send?journal=csu_social").get_data(as_text=True)

                self.assertIn("#1a5276", innovation_html)
                self.assertIn("双创被引论文", innovation_html)
                self.assertNotIn("社科被引论文", innovation_html)
                self.assertIn("#6f4e37", social_html)
                self.assertIn("社科被引论文", social_html)
                self.assertNotIn("双创被引论文", social_html)
        finally:
            app_module.BASE_DIR = old_base
            app_module.RUNTIME_DIR = old_runtime

    def test_runtime_paths_use_user_data_dir_when_configured(self):
        old_runtime = app_module.RUNTIME_DIR
        try:
            with tempfile.TemporaryDirectory() as tmp:
                runtime_dir = Path(tmp)
                app_module.RUNTIME_DIR = runtime_dir

                journal = app_module.JOURNALS["innovation"]
                paths = app_module.build_journal_paths(journal, {"paths": {}})
                app_module.ensure_journal_dirs(paths)

                self.assertEqual(app_module.runtime_path("data/innovation/input"), runtime_dir / "data" / "innovation" / "input")
                self.assertTrue((runtime_dir / "data" / "innovation" / "input").exists())
        finally:
            app_module.RUNTIME_DIR = old_runtime

    def test_cli_send_passes_input_dir_for_original_word_attachments(self):
        config = {
            "paths": {
                "authors_emails_csv": "data/authors_emails.csv",
                "sent_log_csv": "data/sent_log.csv",
                "input_dir": "data/input",
            },
            "email": {"template_file": "templates/notification.html"},
            "journal": {"name": "创新与创业教育"},
            "smtp": {"username": "sender@example.com"},
        }
        recipient = {
            "email": "author@example.com",
            "author_name": "张三",
            "citations": [{"citing_paper_file": "source.docx"}],
        }

        with (
            tempfile.TemporaryDirectory() as tmp,
            patch.object(main_module.os.path, "exists", return_value=True),
            patch.object(main_module, "prepare_email_data", return_value=[recipient]),
            patch.object(main_module, "send_emails") as send_mock,
            patch("builtins.input", return_value="yes"),
        ):
            old_cwd = Path.cwd()
            try:
                main_module.os.chdir(tmp)
                main_module.cmd_send(config)
            finally:
                main_module.os.chdir(old_cwd)

        send_mock.assert_called_once()
        self.assertEqual(send_mock.call_args.kwargs["input_dir"], "data/input")


if __name__ == "__main__":
    unittest.main()
