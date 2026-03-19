"""
第三阶段：邮件的个性化组装与分发
"""

import csv
import os
import smtplib
import ssl
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader


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
            grouped[email] = {
                "email": email,
                "author_name": str(row.get("corresponding_author", "")).strip() or str(row.get("authors", "")).strip(),
                "citations": [],
            }

        grouped[email]["citations"].append({
            "cited_paper_title": str(row.get("title", "")).strip(),
            "cited_paper_authors": str(row.get("authors", "")).strip(),
            "citing_paper_title": str(row.get("article_title", "")).strip(),
            "citing_paper_file": str(row.get("source_file", "")).strip(),
            "cited_paper_doi": str(row.get("doi", "")).strip(),
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
        journal_website=journal_config.get("website", ""),
        year=datetime.now().year,
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
):
    """发送邮件"""
    interval = email_config.get("send_interval_seconds", 8)

    # 加载已发送记录，避免重复
    sent_set = _load_sent_log(sent_log_csv)

    # 建立 SMTP 连接
    server = _connect_smtp(smtp_config)
    if server is None:
        return

    sent_count = 0
    skip_count = 0
    fail_count = 0

    try:
        for i, recipient in enumerate(recipients):
            email_addr = recipient["email"]

            # 检查是否已发送过
            log_key = _make_log_key(recipient)
            if log_key in sent_set:
                print(f"  [{i + 1}/{len(recipients)}] 跳过（已发送）: {email_addr}")
                skip_count += 1
                continue

            subject, html_body = render_email(recipient, template_file, journal_config)

            try:
                msg = MIMEMultipart("alternative")
                msg["From"] = f"{smtp_config.get('sender_name', '')} <{smtp_config['username']}>"
                msg["To"] = email_addr
                msg["Subject"] = subject

                msg.attach(MIMEText(html_body, "html", "utf-8"))

                server.send_message(msg)
                sent_count += 1
                print(f"  [{i + 1}/{len(recipients)}] ✓ 已发送: {email_addr}")

                _append_sent_log(sent_log_csv, recipient, "成功")
                sent_set.add(log_key)

            except Exception as e:
                fail_count += 1
                print(f"  [{i + 1}/{len(recipients)}] ✗ 发送失败: {email_addr} — {e}")
                _append_sent_log(sent_log_csv, recipient, f"失败: {e}")

            # 间隔发送
            if i < len(recipients) - 1:
                time.sleep(interval)

    finally:
        server.quit()

    print(f"\n发送完成: 成功 {sent_count}, 跳过 {skip_count}, 失败 {fail_count}")


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
    """生成去重用的 key"""
    titles = sorted(c["cited_paper_title"] for c in recipient["citations"])
    return f"{recipient['email']}|{'|'.join(titles)}"


def _load_sent_log(sent_log_csv: str) -> set:
    """加载已发送记录"""
    sent_set = set()
    if not os.path.exists(sent_log_csv):
        return sent_set

    with open(sent_log_csv, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") == "成功":
                sent_set.add(row.get("log_key", ""))

    return sent_set


def _append_sent_log(sent_log_csv: str, recipient: dict, status: str):
    """追加发送记录到日志"""
    os.makedirs(os.path.dirname(sent_log_csv), exist_ok=True)

    file_exists = os.path.exists(sent_log_csv)

    with open(sent_log_csv, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "timestamp", "email", "author_name", "citation_count", "status", "log_key",
        ])
        if not file_exists:
            writer.writeheader()

        writer.writerow({
            "timestamp": datetime.now().isoformat(),
            "email": recipient["email"],
            "author_name": recipient["author_name"],
            "citation_count": len(recipient["citations"]),
            "status": status,
            "log_key": _make_log_key(recipient),
        })
