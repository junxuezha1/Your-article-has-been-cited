"""
《创新与创业教育》引用通知系统 — 主入口

使用方式:
    python main.py web         — 启动 Web 网页界面（推荐）
    python main.py extract     — 命令行：提取参考文献
    python main.py lookup      — 命令行：检索被引作者邮箱
    python main.py merge       — 命令行：合并人工补充的邮箱数据
    python main.py preview     — 命令行：生成邮件预览
    python main.py send        — 命令行：正式发送邮件
"""

import os
import sys

import yaml

from extract_references import process_input_directory, save_references_csv
from lookup_authors import lookup_emails, merge_manual_supplement, save_results
from send_emails import prepare_email_data, preview_emails, send_emails


def load_config(config_path: str = "config.yaml") -> dict:
    if not os.path.exists(config_path):
        print(f"错误：配置文件不存在 — {config_path}")
        print("请复制 config.yaml 并填写您的 SMTP 和 API 配置")
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def cmd_web():
    """启动 Web 界面"""
    from app import app
    app.run(debug=True, host="127.0.0.1", port=5000)


def cmd_extract(config: dict):
    print("=" * 60)
    print("第一阶段：从文章中提取参考文献")
    print("=" * 60)
    paths = config["paths"]
    results = process_input_directory(paths["input_dir"])
    if not results:
        print("未找到任何文件，请将 Word/PDF 文件放入 data/input/ 目录")
        return
    save_references_csv(results, paths["references_csv"])


def cmd_lookup(config: dict):
    print("=" * 60)
    print("第二阶段：检索被引作者邮箱")
    print("=" * 60)
    paths = config["paths"]
    ref_csv = paths["references_csv"]
    if not os.path.exists(ref_csv):
        print(f"错误：参考文献数据文件不存在 — {ref_csv}")
        print("请先运行: python main.py extract")
        return
    df = lookup_emails(ref_csv, config)
    save_results(df, paths["authors_emails_csv"], paths["failed_lookup_csv"])


def cmd_merge(config: dict):
    print("=" * 60)
    print("合并人工补充的邮箱数据")
    print("=" * 60)
    paths = config["paths"]
    merge_manual_supplement(paths["authors_emails_csv"], paths["manual_supplement_csv"])


def cmd_preview(config: dict):
    print("=" * 60)
    print("第三阶段（预览）：生成邮件内容供审核")
    print("=" * 60)
    paths = config["paths"]
    emails_csv = paths["authors_emails_csv"]
    if not os.path.exists(emails_csv):
        print(f"错误：邮箱数据文件不存在 — {emails_csv}")
        print("请先运行: python main.py lookup")
        return
    recipients = prepare_email_data(emails_csv)
    if not recipients:
        print("没有可发送的收件人")
        return
    preview_emails(
        recipients,
        config["email"]["template_file"],
        config["journal"],
        paths.get("preview_dir", "data/preview"),
    )


def cmd_send(config: dict):
    print("=" * 60)
    print("第三阶段（发送）：正式发送引用通知邮件")
    print("=" * 60)
    paths = config["paths"]
    emails_csv = paths["authors_emails_csv"]
    if not os.path.exists(emails_csv):
        print(f"错误：邮箱数据文件不存在 — {emails_csv}")
        print("请先运行: python main.py lookup")
        return
    recipients = prepare_email_data(emails_csv)
    if not recipients:
        print("没有可发送的收件人")
        return
    print(f"\n即将向 {len(recipients)} 位作者发送引用通知邮件。")
    confirm = input("确认发送？(输入 yes 继续): ").strip().lower()
    if confirm != "yes":
        print("已取消发送")
        return
    send_emails(
        recipients,
        config["email"]["template_file"],
        config["journal"],
        config["smtp"],
        config["email"],
        paths["sent_log_csv"],
    )


def main():
    if len(sys.argv) < 2:
        # 默认启动 Web 界面
        cmd_web()
        return

    command = sys.argv[1].lower()

    if command == "web":
        cmd_web()
        return

    config = load_config()

    commands = {
        "extract": cmd_extract,
        "lookup": cmd_lookup,
        "merge": cmd_merge,
        "preview": cmd_preview,
        "send": cmd_send,
    }

    if command == "all":
        cmd_extract(config)
        print()
        cmd_lookup(config)
        print()
        cmd_preview(config)
        print("\n全部完成！请检查 data/preview/ 中的邮件预览，确认无误后运行:")
        print("  python main.py send")
    elif command in commands:
        commands[command](config)
    else:
        print(f"未知命令: {command}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
