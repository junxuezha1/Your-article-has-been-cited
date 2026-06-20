"""
第三阶段：邮件的个性化组装与分发
"""

import csv
import os
import smtplib
import ssl
import time
from datetime import datetime
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

import pandas as pd
from jinja2 import Environment, FileSystemLoader


SENT_LOG_FIELDS = [
    "timestamp",
    "email",
    "author_name",
    "cited_paper_title",
    "citing_paper_title",
    "status",
]

LEGACY_SENT_LOG_FIELDS = [
    "timestamp",
    "email",
    "author_name",
    "citation_count",
    "status",
    "log_key",
]


def _parse_legacy_log_key(log_key: str) -> str:
    """从旧版 log_key 中尽量恢复被引用文章标题。"""
    text = str(log_key).strip()
    if not text or "|" not in text:
        return ""

    _, _, payload = text.partition("|")
    if not payload:
        return ""

    first_item = payload.split("|", 1)[0].strip()
    if "@" in first_item:
        first_item = first_item.rsplit("@", 1)[0]
    return first_item.strip()


def normalize_sent_log_df(df: pd.DataFrame) -> pd.DataFrame:
    """统一发送日志字段，兼容旧版与新版 CSV 结构。"""
    if df is None or df.empty:
        return pd.DataFrame(columns=SENT_LOG_FIELDS)

    normalized = df.copy().fillna("")

    if set(SENT_LOG_FIELDS).issubset(normalized.columns):
        return normalized.reindex(columns=SENT_LOG_FIELDS, fill_value="")

    if set(LEGACY_SENT_LOG_FIELDS).issubset(normalized.columns):
        normalized["cited_paper_title"] = normalized["log_key"].apply(_parse_legacy_log_key)
        # 旧日志不含“引用了的文章”标题，这里保留空值，避免写入误导性数据。
        normalized["citing_paper_title"] = ""
        return normalized.reindex(columns=SENT_LOG_FIELDS, fill_value="")

    for field in SENT_LOG_FIELDS:
        if field not in normalized.columns:
            normalized[field] = ""
    return normalized.reindex(columns=SENT_LOG_FIELDS, fill_value="")


def load_sent_log(sent_log_csv: str) -> pd.DataFrame:
    """读取发送日志并标准化字段。"""
    if not os.path.exists(sent_log_csv):
        return pd.DataFrame(columns=SENT_LOG_FIELDS)

    try:
        df = pd.read_csv(sent_log_csv, encoding="utf-8-sig").fillna("")
    except pd.errors.EmptyDataError:
        return pd.DataFrame(columns=SENT_LOG_FIELDS)

    return normalize_sent_log_df(df)


def _ensure_sent_log_schema(sent_log_csv: str):
    """若发现旧版发送日志格式，则先迁移为新版字段后再追加写入。"""
    if not os.path.exists(sent_log_csv):
        return

    try:
        df = pd.read_csv(sent_log_csv, encoding="utf-8-sig").fillna("")
    except pd.errors.EmptyDataError:
        return

    normalized = normalize_sent_log_df(df)
    if list(df.columns) != SENT_LOG_FIELDS:
        normalized.to_csv(sent_log_csv, index=False, encoding="utf-8-sig")


def prepare_email_data(emails_csv: str) -> list[dict]:
    """
    读取邮箱数据，按收件人邮箱聚合：
    同一作者被多篇文章引用时，合并为一封邮件。
    """
    df = pd.read_csv(emails_csv, encoding="utf-8-sig")

    # 按邮箱分组
    grouped = {}
    for _, row in df.iterrows():
        email = str(row.get("email", "")).strip()
        if not email or email == "nan":
            continue

        if email not in grouped:
            def _clean(val):
                s = str(val).strip()
                return "" if s == "nan" else s

            author_name = (
                _clean(row.get("matched_name", ""))
                or _clean(row.get("corresponding_author", ""))
                or _clean(row.get("authors", ""))
            )
            grouped[email] = {
                "email": email,
                "author_name": author_name,
                "citations": [],
            }

        grouped[email]["citations"].append({
            "cited_paper_title": str(row.get("title", "")).strip(),
            "cited_paper_authors": str(row.get("authors", "")).strip(),
            "citing_paper_title": str(row.get("article_title", "")).strip(),
            "citing_paper_authors": str(row.get("article_authors", "")).strip(),
            "citing_paper_file": str(row.get("source_file", "")).strip(),
        })

    recipients = list(grouped.values())
    print(f"共 {len(recipients)} 位收件人（{sum(len(r['citations']) for r in recipients)} 条引用记录）")
    return recipients


def render_email(recipient: dict, template_file: str, journal_config: dict) -> tuple[str, str]:
    """渲染单封邮件的主题和正文"""
    template_dir = os.path.dirname(template_file)
    template_name = os.path.basename(template_file)

    env = Environment(loader=FileSystemLoader(template_dir), autoescape=True)
    template = env.get_template(template_name)

    html_body = template.render(
        author_name=recipient["author_name"],
        citations=recipient["citations"],
        journal_name=journal_config.get("name", "创新与创业教育"),
        journal_name_en=journal_config.get("name_en", "Innovation and Entrepreneurship Education"),
        journal_intro=journal_config.get("intro", ""),
        journal_website=journal_config.get("website", ""),
        journal_theme_color=journal_config.get("theme_color", "#1a5276"),
        journal_theme_bg=journal_config.get("theme_bg", "#f4f7fb"),
        editorial_office=journal_config.get("editorial_office", f"《{journal_config.get('name', '创新与创业教育')}》编辑部"),
        attachment_note=journal_config.get(
            "attachment_note",
            "我们在附件中给您发送了原文的word版本，如您在后续大作中引用，请以知网PDF版本为准。",
        ),
        year=datetime.now().year,
        send_date=datetime.now().strftime("%Y年%m月%d日"),
    )

    subject = f"您的文章在《{journal_config.get('name', '创新与创业教育')}》中被引用的通知"
    return subject, html_body


def preview_emails(recipients: list[dict], template_file: str, journal_config: dict, preview_dir: str):
    """生成邮件预览文件，供人工审核"""
    os.makedirs(preview_dir, exist_ok=True)

    for i, recipient in enumerate(recipients):
        subject, html_body = render_email(recipient, template_file, journal_config)

        filename = f"preview_{i + 1}_{recipient['email'].replace('@', '_at_')}.html"
        filepath = os.path.join(preview_dir, filename)

        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"<!-- 收件人: {recipient['email']} -->\n")
            f.write(f"<!-- 主题: {subject} -->\n")
            f.write(f"<!-- 引用数: {len(recipient['citations'])} -->\n\n")
            f.write(html_body)

    print(f"\n邮件预览已生成到: {preview_dir}")
    print(f"共 {len(recipients)} 封邮件，请逐一检查后再执行发送")


def send_emails(
    recipients: list[dict],
    template_file: str,
    journal_config: dict,
    smtp_config: dict,
    email_config: dict,
    sent_log_csv: str,
    html_overrides: dict = None,
    input_dir: str = "",
):
    """发送邮件"""
    interval = email_config.get("send_interval_seconds", 8)
    if html_overrides is None:
        html_overrides = {}

    # 建立 SMTP 连接
    server = _connect_smtp(smtp_config)
    if server is None:
        return

    sent_count = 0
    fail_count = 0

    try:
        for i, recipient in enumerate(recipients):
            email_addr = recipient["email"]

            subject, html_body = render_email(recipient, template_file, journal_config)
            if email_addr in html_overrides:
                html_body = html_overrides[email_addr]

            try:
                msg = MIMEMultipart("mixed")
                msg["From"] = formataddr((str(Header(smtp_config.get("sender_name", ""), "utf-8")), smtp_config["username"]))
                msg["To"] = email_addr
                msg["Subject"] = subject

                alt_part = MIMEMultipart("alternative")
                alt_part.attach(MIMEText(html_body, "html", "utf-8"))
                msg.attach(alt_part)

                if input_dir:
                    seen_files = set()
                    for citation in recipient["citations"]:
                        fname = citation.get("citing_paper_file", "")
                        if fname and fname not in seen_files:
                            fpath = os.path.join(input_dir, fname)
                            if os.path.exists(fpath):
                                with open(fpath, "rb") as f:
                                    part = MIMEBase("application", "octet-stream")
                                    part.set_payload(f.read())
                                encoders.encode_base64(part)
                                part.add_header(
                                    "Content-Disposition", "attachment",
                                    filename=Header(fname, "utf-8").encode()
                                )
                                msg.attach(part)
                            seen_files.add(fname)

                server.send_message(msg)
                sent_count += 1
                print(f"  [{i + 1}/{len(recipients)}] ✓ 已发送: {email_addr}")

                _append_sent_log(sent_log_csv, recipient, "成功")

            except Exception as e:
                fail_count += 1
                print(f"  [{i + 1}/{len(recipients)}] ✗ 发送失败: {email_addr} — {e}")
                _append_sent_log(sent_log_csv, recipient, f"失败: {e}")

            # 间隔发送
            if i < len(recipients) - 1:
                time.sleep(interval)

    finally:
        try:
            server.quit()
        except Exception:
            pass

    print(f"\n发送完成: 成功 {sent_count}, 失败 {fail_count}")


def _connect_smtp(smtp_config: dict):
    """建立 SMTP 连接"""
    server_addr = smtp_config["server"]
    port = smtp_config["port"]
    username = smtp_config["username"]
    password = smtp_config["password"]
    use_ssl = smtp_config.get("use_ssl", True)

    try:
        if use_ssl:
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(server_addr, port, context=context)
        else:
            server = smtplib.SMTP(server_addr, port)
            server.starttls()

        server.login(username, password)
        print(f"✓ SMTP 连接成功: {server_addr}:{port}")
        return server

    except Exception as e:
        print(f"✗ SMTP 连接失败: {e}")
        print("  请检查 config.yaml 中的 SMTP 配置:")
        print(f"    服务器: {server_addr}, 端口: {port}")
        print(f"    用户名: {username}")
        print("    密码/授权码是否正确")
        return None


def _make_log_key(recipient: dict) -> str:
    """生成去重用的 key，包含引用来源文章避免跨期次误判"""
    pairs = sorted(
        f"{c['cited_paper_title']}@{c['citing_paper_file']}"
        for c in recipient["citations"]
    )
    return f"{recipient['email']}|{'|'.join(pairs)}"


def _append_sent_log(sent_log_csv: str, recipient: dict, status: str):
    """追加发送记录到日志（每条引用写一行，便于 Excel 导出详细信息）"""
    log_dir = os.path.dirname(sent_log_csv)
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)

    file_exists = os.path.exists(sent_log_csv)
    if file_exists:
        _ensure_sent_log_schema(sent_log_csv)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(sent_log_csv, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=SENT_LOG_FIELDS)
        if not file_exists:
            writer.writeheader()

        citations = recipient.get("citations") or []
        if citations:
            for c in citations:
                writer.writerow({
                    "timestamp": timestamp,
                    "email": recipient["email"],
                    "author_name": recipient["author_name"],
                    "cited_paper_title": c.get("cited_paper_title", ""),
                    "citing_paper_title": c.get("citing_paper_title", ""),
                    "status": status,
                })
        else:
            writer.writerow({
                "timestamp": timestamp,
                "email": recipient["email"],
                "author_name": recipient["author_name"],
                "cited_paper_title": "",
                "citing_paper_title": "",
                "status": status,
            })
