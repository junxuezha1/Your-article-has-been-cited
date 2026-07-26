import tempfile
import unittest
from email.header import decode_header
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import app as app_module
import main as main_module
from extract_references import split_reference_rows
from send_emails import prepare_email_data, render_email, repair_citing_authors_from_input, send_emails


class MultiJournalBehaviorTests(unittest.TestCase):
    @staticmethod
    def _decode_mime_header(value):
        parts = []
        for chunk, encoding in decode_header(value):
            if isinstance(chunk, bytes):
                parts.append(chunk.decode(encoding or "utf-8"))
            else:
                parts.append(chunk)
        return "".join(parts)

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
            "show_name_en_in_email": False,
            "intro": "综合性人文社科学术期刊。",
            "editorial_office": "《中南大学学报（社会科学版）》编辑部",
        }

        subject, html = render_email(recipient, "templates/notification.html", journal)

        self.assertIn("中南大学学报（社会科学版）", subject)
        self.assertIn("综合性人文社科学术期刊", html)
        self.assertIn("《中南大学学报（社会科学版）》编辑部", html)
        self.assertNotIn("Journal of Central South University", html)

    def test_email_content_places_formal_notice_before_citations_and_intro_last(self):
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
            "name": "创新与创业教育",
            "name_en": "Innovation and Entrepreneurship Education",
            "show_name_en_in_email": False,
            "intro": "《创新与创业教育》是教育部主管、中南大学主办的学术期刊，创刊于2010年，为中国人文社会科学 AMI 综合评价 A 刊核心期刊扩展版，被 CNKI、万方数据、维普资讯等收录。本刊常设“学术前沿”“教学研究”“教育治理”“数智教育”“劳动教育”“创业管理”“双创实践”“技术创新”“青年论坛”等栏目，致力于打造集学术引领、实践创新与政策研讨于一体的高水平学术交流平台。",
            "editorial_office": "《创新与创业教育》编辑部",
            "attachment_note": "以上引用信息供您知悉与参考，具体内容请以本刊正式出版版本为准。",
        }

        _, html = render_email(recipient, "templates/notification.html", journal)

        notice_idx = html.index("很荣幸地通知您：本刊刊发文章引用了您的研究成果。有关信息如下：")
        citation_idx = html.index("引用该文献的本刊文章")
        intro_idx = html.index("中国人文社会科学 AMI 综合评价 A 刊核心期刊扩展版")
        salutation_idx = html.index("此致")

        self.assertLess(notice_idx, citation_idx)
        self.assertLess(citation_idx, intro_idx)
        self.assertLess(intro_idx, salutation_idx)
        self.assertIn("感谢您的研究成果对相关学术领域的贡献，也感谢您的成果为本刊作者提供了重要参考", html)
        self.assertIn("以上引用信息供您知悉与参考，具体内容请以本刊正式出版版本为准。", html)
        self.assertIn("本刊常设“学术前沿”“教学研究”“教育治理”“数智教育”“劳动教育”“创业管理”“双创实践”“技术创新”“青年论坛”等栏目", html)
        self.assertNotIn("Innovation and Entrepreneurship Education", html)
        self.assertIn("《来源论文》", html)
        self.assertIn("《被引论文》", html)

    def test_email_content_replaces_nested_title_marks(self):
        recipient = {
            "email": "author@example.com",
            "author_name": "张三",
            "citations": [
                {
                    "cited_paper_title": "《资本论》中的合作思想研究",
                    "cited_paper_authors": "张三",
                    "citing_paper_title": "论《资本论》与现代企业制度",
                    "citing_paper_authors": "李四",
                    "citing_paper_file": "source.docx",
                }
            ],
        }

        _, html = render_email(recipient, "templates/notification.html", {"name": "创新与创业教育"})

        self.assertIn("《论〈资本论〉与现代企业制度》", html)
        self.assertIn("《〈资本论〉中的合作思想研究》", html)
        self.assertNotIn("《论《资本论》与现代企业制度》", html)
        self.assertNotIn("《《资本论》中的合作思想研究》", html)

    def test_prepare_email_data_cleans_nan_citation_fields(self):
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "authors_emails.csv"
            pd.DataFrame([
                {
                    "email": "author@example.com",
                    "matched_name": "聂辉华",
                    "title": "保护市场的联邦主义及其批判",
                    "authors": "杨其静, 聂辉华",
                    "article_title": "央地财政分权何以影响区域协调发展？——一个政治经济学分析",
                    "article_authors": None,
                    "source_file": "source.docx",
                }
            ]).to_csv(csv_path, index=False, encoding="utf-8-sig")

            recipients = prepare_email_data(str(csv_path))

        self.assertEqual(recipients[0]["citations"][0]["citing_paper_authors"], "")

    def test_repair_citing_authors_replaces_title_fragment_from_source_file(self):
        recipients = [
            {
                "email": "author@example.com",
                "author_name": "张茂杰",
                "citations": [
                    {
                        "cited_paper_title": "新时代主流意识形态的情感叙事策略审思",
                        "cited_paper_authors": "肖唤元, 张茂杰",
                        "citing_paper_title": "数字时代的意识形态叙事： 价值意蕴、潜在危机与化解策略",
                        "citing_paper_authors": "价值意蕴、潜在危机与化解策略",
                        "citing_paper_file": "source.doc",
                    }
                ],
            }
        ]

        with tempfile.TemporaryDirectory() as tmp, patch(
            "send_emails._extract_source_metadata",
            return_value={"article_authors": "曹清燕，欧露雯"},
        ):
            (Path(tmp) / "source.doc").write_bytes(b"fake")
            repair_citing_authors_from_input(recipients, tmp)

        self.assertEqual(recipients[0]["citations"][0]["citing_paper_authors"], "曹清燕，欧露雯")

    def test_repair_citing_authors_fills_blank_from_source_file(self):
        recipients = [
            {
                "email": "author@example.com",
                "author_name": "陈婉玲",
                "citations": [
                    {
                        "cited_paper_title": "区域发展权的权利结构与实现路径",
                        "cited_paper_authors": "陈婉玲, 周浩然",
                        "citing_paper_title": "央地财政分权何以影响区域协调发展？——一个政治经济学分析",
                        "citing_paper_authors": "",
                        "citing_paper_file": "source.doc",
                    }
                ],
            }
        ]

        with tempfile.TemporaryDirectory() as tmp, patch(
            "send_emails._extract_source_metadata",
            return_value={"article_authors": "肖芸"},
        ):
            (Path(tmp) / "source.doc").write_bytes(b"fake")
            repair_citing_authors_from_input(recipients, tmp)

        self.assertEqual(recipients[0]["citations"][0]["citing_paper_authors"], "肖芸")

    def test_journal_defaults_match_requested_introductions(self):
        innovation = app_module.JOURNALS["innovation"]
        social = app_module.JOURNALS["csu_social"]

        self.assertFalse(innovation["show_name_en_in_email"])
        self.assertEqual(
            innovation["attachment_note"],
            "以上引用信息供您知悉与参考，具体内容请以本刊正式出版版本为准。",
        )
        self.assertIn("本刊常设“学术前沿”“教学研究”“教育治理”“数智教育”“劳动教育”“创业管理”“双创实践”“技术创新”“青年论坛”等栏目", innovation["intro"])
        self.assertNotIn("创刊于1995年，双月刊，单月出版", social["intro"])

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

    def test_send_emails_attaches_original_word_file_when_input_dir_is_available(self):
        class FakeSMTP:
            def __init__(self):
                self.messages = []
                self.quit_called = False

            def send_message(self, msg):
                self.messages.append(msg)

            def quit(self):
                self.quit_called = True

        fake_server = FakeSMTP()
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

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            input_dir = tmp_path / "input"
            input_dir.mkdir()
            (input_dir / "source.docx").write_bytes(b"fake word content")

            with patch("send_emails._connect_smtp", return_value=fake_server):
                send_emails(
                    [recipient],
                    "templates/notification.html",
                    {"name": "创新与创业教育"},
                    {"username": "sender@example.com", "sender_name": "编辑部"},
                    {"send_interval_seconds": 0},
                    str(tmp_path / "sent_log.csv"),
                    input_dir=str(input_dir),
                )

        self.assertEqual(len(fake_server.messages), 1)
        filenames = [
            self._decode_mime_header(part.get_filename())
            for part in fake_server.messages[0].walk()
            if part.get_filename()
        ]
        self.assertIn("source.docx", filenames)
        self.assertTrue(fake_server.quit_called)


if __name__ == "__main__":
    unittest.main()
