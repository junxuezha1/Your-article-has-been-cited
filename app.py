"""
《创新与创业教育》引用通知系统 — Flask Web 应用
"""

import os
import json
import threading
from pathlib import Path

import yaml
import pandas as pd
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, jsonify, send_from_directory,
)
from werkzeug.utils import secure_filename

from extract_references import process_input_directory, save_references_csv
from lookup_authors import lookup_emails, save_results, merge_manual_supplement
from send_emails import prepare_email_data, render_email, send_emails

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.yaml"
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
PREVIEW_DIR = DATA_DIR / "preview"

app = Flask(__name__)
app.secret_key = "citation-notifier-secret-key-change-me"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

ALLOWED_EXTENSIONS = {".doc", ".docx", ".pdf"}

# ---- 全局任务状态 ----
task_status = {
    "running": False,
    "stage": "",
    "progress": "",
    "log": [],
}


def load_config() -> dict:
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    return {}


def save_config(config: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


def get_data_stats() -> dict:
    """获取各阶段数据统计"""
    stats = {
        "input_files": 0,
        "references": 0,
        "emails_found": 0,
        "emails_failed": 0,
        "emails_sent": 0,
    }

    # 输入文件数
    if INPUT_DIR.exists():
        stats["input_files"] = len([
            f for f in INPUT_DIR.iterdir()
            if f.suffix.lower() in ALLOWED_EXTENSIONS
        ])

    config = load_config()
    paths = config.get("paths", {})

    # 参考文献数
    ref_csv = BASE_DIR / paths.get("references_csv", "data/references.csv")
    if ref_csv.exists():
        try:
            df = pd.read_csv(ref_csv, encoding="utf-8-sig").fillna("")
            stats["references"] = len(df)
        except Exception:
            pass

    # 邮箱检索结果
    emails_csv = BASE_DIR / paths.get("authors_emails_csv", "data/authors_emails.csv")
    if emails_csv.exists():
        try:
            df = pd.read_csv(emails_csv, encoding="utf-8-sig").fillna("")
            stats["emails_found"] = len(df)
        except Exception:
            pass

    failed_csv = BASE_DIR / paths.get("failed_lookup_csv", "data/failed_lookup.csv")
    if failed_csv.exists():
        try:
            df = pd.read_csv(failed_csv, encoding="utf-8-sig").fillna("")
            stats["emails_failed"] = len(df)
        except Exception:
            pass

    # 已发送
    sent_csv = BASE_DIR / paths.get("sent_log_csv", "data/sent_log.csv")
    if sent_csv.exists():
        try:
            df = pd.read_csv(sent_csv, encoding="utf-8-sig").fillna("")
            stats["emails_sent"] = len(df[df["status"] == "成功"])
        except Exception:
            pass

    return stats


# ==================== 路由 ====================

@app.route("/")
def index():
    stats = get_data_stats()
    return render_template("index.html", stats=stats)


# ---- 配置页 ----

@app.route("/config", methods=["GET", "POST"])
def config_page():
    config = load_config()

    if request.method == "POST":
        # 从表单更新配置
        config.setdefault("smtp", {})
        config["smtp"]["server"] = request.form.get("smtp_server", "").strip()
        config["smtp"]["port"] = int(request.form.get("smtp_port", 465))
        config["smtp"]["use_ssl"] = request.form.get("smtp_ssl") == "on"
        config["smtp"]["username"] = request.form.get("smtp_username", "").strip()
        config["smtp"]["password"] = request.form.get("smtp_password", "").strip()
        config["smtp"]["sender_name"] = request.form.get("smtp_sender_name", "").strip()

        config.setdefault("crossref", {})
        config["crossref"]["mailto"] = request.form.get("crossref_mailto", "").strip()

        config.setdefault("journal", {})
        config["journal"]["name"] = request.form.get("journal_name", "").strip()
        config["journal"]["name_en"] = request.form.get("journal_name_en", "").strip()
        config["journal"]["website"] = request.form.get("journal_website", "").strip()

        save_config(config)
        flash("配置已保存", "success")
        return redirect(url_for("config_page"))

    return render_template("config_page.html", config=config)


@app.route("/config/test-smtp", methods=["POST"])
def test_smtp():
    """测试 SMTP 连接"""
    import smtplib
    import ssl

    config = load_config()
    smtp = config.get("smtp", {})

    try:
        if smtp.get("use_ssl", True):
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(smtp["server"], smtp["port"], context=context, timeout=10)
        else:
            server = smtplib.SMTP(smtp["server"], smtp["port"], timeout=10)
            server.starttls()

        server.login(smtp["username"], smtp["password"])
        server.quit()
        return jsonify({"success": True, "message": "SMTP 连接成功！"})
    except Exception as e:
        return jsonify({"success": False, "message": f"连接失败: {e}"})


# ---- 第一阶段：上传与提取 ----

@app.route("/extract")
def extract_page():
    # 列出已上传的文件
    files = []
    if INPUT_DIR.exists():
        for f in sorted(INPUT_DIR.iterdir()):
            if f.suffix.lower() in ALLOWED_EXTENSIONS:
                size_kb = f.stat().st_size / 1024
                files.append({"name": f.name, "size": f"{size_kb:.1f} KB", "type": f.suffix})

    # 读取已提取的参考文献
    config = load_config()
    ref_csv = BASE_DIR / config.get("paths", {}).get("references_csv", "data/references.csv")
    references = []
    if ref_csv.exists():
        try:
            df = pd.read_csv(ref_csv, encoding="utf-8-sig").fillna("").fillna("")
            references = df.to_dict("records")
        except Exception:
            pass

    return render_template("extract.html", files=files, references=references)


@app.route("/extract/upload", methods=["POST"])
def upload_files():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    uploaded = request.files.getlist("files")
    count = 0
    for f in uploaded:
        if f.filename:
            ext = Path(f.filename).suffix.lower()
            if ext in ALLOWED_EXTENSIONS:
                filename = secure_filename(f.filename)
                # 保留中文文件名
                if not filename or filename == ext:
                    filename = f.filename
                f.save(str(INPUT_DIR / filename))
                count += 1

    flash(f"已上传 {count} 个文件", "success")
    return redirect(url_for("extract_page"))


@app.route("/extract/delete/<filename>", methods=["POST"])
def delete_file(filename):
    filepath = INPUT_DIR / filename
    if filepath.exists():
        filepath.unlink()
        flash(f"已删除: {filename}", "info")
    return redirect(url_for("extract_page"))


@app.route("/extract/run", methods=["POST"])
def run_extract():
    config = load_config()
    paths = config.get("paths", {})

    results = process_input_directory(str(INPUT_DIR))
    if not results:
        flash("未找到任何可处理的文件", "warning")
        return redirect(url_for("extract_page"))

    ref_csv = str(BASE_DIR / paths.get("references_csv", "data/references.csv"))
    rows = save_references_csv(results, ref_csv)

    total_refs = sum(len(r["references"]) for r in results)
    errors = [r for r in results if r.get("error")]

    msg = f"提取完成：{len(results)} 篇文章，共 {total_refs} 条参考文献"
    if errors:
        msg += f"，{len(errors)} 篇有警告"

    flash(msg, "success")
    return redirect(url_for("extract_page"))


# ---- 第二阶段：邮箱检索 ----

@app.route("/lookup")
def lookup_page():
    config = load_config()
    paths = config.get("paths", {})

    # 检索结果
    emails_data = []
    emails_csv = BASE_DIR / paths.get("authors_emails_csv", "data/authors_emails.csv")
    if emails_csv.exists():
        try:
            df = pd.read_csv(emails_csv, encoding="utf-8-sig").fillna("")
            emails_data = df.to_dict("records")
        except Exception:
            pass

    # 失败记录
    failed_data = []
    failed_csv = BASE_DIR / paths.get("failed_lookup_csv", "data/failed_lookup.csv")
    if failed_csv.exists():
        try:
            df = pd.read_csv(failed_csv, encoding="utf-8-sig").fillna("")
            failed_data = df.to_dict("records")
        except Exception:
            pass

    # 检查是否有参考文献数据
    ref_csv = BASE_DIR / paths.get("references_csv", "data/references.csv")
    has_references = ref_csv.exists()

    return render_template(
        "lookup.html",
        emails_data=emails_data,
        failed_data=failed_data,
        has_references=has_references,
        task_status=task_status,
    )


@app.route("/lookup/run", methods=["POST"])
def run_lookup():
    if task_status["running"]:
        flash("已有任务正在运行，请等待完成", "warning")
        return redirect(url_for("lookup_page"))

    config = load_config()
    paths = config.get("paths", {})
    ref_csv = str(BASE_DIR / paths.get("references_csv", "data/references.csv"))

    if not os.path.exists(ref_csv):
        flash("请先执行第一阶段（提取参考文献）", "warning")
        return redirect(url_for("lookup_page"))

    # 在后台线程中运行检索（避免超时）
    def run_in_background():
        task_status["running"] = True
        task_status["stage"] = "邮箱检索"
        task_status["log"] = []
        try:
            df = lookup_emails(ref_csv, config)
            emails_csv = str(BASE_DIR / paths.get("authors_emails_csv", "data/authors_emails.csv"))
            failed_csv = str(BASE_DIR / paths.get("failed_lookup_csv", "data/failed_lookup.csv"))
            save_results(df, emails_csv, failed_csv)
            task_status["progress"] = "完成"
        except Exception as e:
            task_status["progress"] = f"出错: {e}"
        finally:
            task_status["running"] = False

    thread = threading.Thread(target=run_in_background, daemon=True)
    thread.start()

    flash("邮箱检索已开始（后台运行中），请稍候刷新页面查看结果", "info")
    return redirect(url_for("lookup_page"))


@app.route("/lookup/upload-supplement", methods=["POST"])
def upload_supplement():
    """上传人工补充的邮箱 CSV"""
    config = load_config()
    paths = config.get("paths", {})

    f = request.files.get("supplement_file")
    if not f or not f.filename:
        flash("请选择文件", "warning")
        return redirect(url_for("lookup_page"))

    supp_csv = str(BASE_DIR / paths.get("manual_supplement_csv", "data/manual_supplement.csv"))
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    f.save(supp_csv)

    emails_csv = str(BASE_DIR / paths.get("authors_emails_csv", "data/authors_emails.csv"))
    merge_manual_supplement(emails_csv, supp_csv)

    flash("人工补充数据已合并", "success")
    return redirect(url_for("lookup_page"))


# ---- 第三阶段：预览与发送 ----

@app.route("/send")
def send_page():
    config = load_config()
    paths = config.get("paths", {})

    emails_csv = BASE_DIR / paths.get("authors_emails_csv", "data/authors_emails.csv")
    recipients = []
    if emails_csv.exists():
        try:
            recipients = prepare_email_data(str(emails_csv))
        except Exception:
            pass

    # 生成预览
    previews = []
    template_file = str(BASE_DIR / config.get("email", {}).get("template_file", "templates/notification.html"))
    journal_config = config.get("journal", {})

    for r in recipients:
        try:
            subject, html_body = render_email(r, template_file, journal_config)
            previews.append({
                "email": r["email"],
                "author_name": r["author_name"],
                "citation_count": len(r["citations"]),
                "subject": subject,
                "html_body": html_body,
            })
        except Exception as e:
            previews.append({
                "email": r["email"],
                "author_name": r["author_name"],
                "citation_count": len(r["citations"]),
                "subject": "渲染失败",
                "html_body": f"<p>模板渲染错误: {e}</p>",
            })

    # 已发送记录
    sent_data = []
    sent_csv = BASE_DIR / paths.get("sent_log_csv", "data/sent_log.csv")
    if sent_csv.exists():
        try:
            df = pd.read_csv(sent_csv, encoding="utf-8-sig").fillna("")
            sent_data = df.to_dict("records")
        except Exception:
            pass

    return render_template(
        "send_page.html",
        previews=previews,
        sent_data=sent_data,
        task_status=task_status,
    )


@app.route("/send/run", methods=["POST"])
def run_send():
    if task_status["running"]:
        flash("已有任务正在运行，请等待完成", "warning")
        return redirect(url_for("send_page"))

    config = load_config()
    paths = config.get("paths", {})

    emails_csv = str(BASE_DIR / paths.get("authors_emails_csv", "data/authors_emails.csv"))
    if not os.path.exists(emails_csv):
        flash("没有可发送的邮箱数据", "warning")
        return redirect(url_for("send_page"))

    recipients = prepare_email_data(emails_csv)
    if not recipients:
        flash("没有可发送的收件人", "warning")
        return redirect(url_for("send_page"))

    template_file = str(BASE_DIR / config.get("email", {}).get("template_file", "templates/notification.html"))
    journal_config = config.get("journal", {})
    smtp_config = config.get("smtp", {})
    email_config = config.get("email", {})
    sent_log_csv = str(BASE_DIR / paths.get("sent_log_csv", "data/sent_log.csv"))

    def run_in_background():
        task_status["running"] = True
        task_status["stage"] = "邮件发送"
        try:
            send_emails(
                recipients, template_file, journal_config,
                smtp_config, email_config, sent_log_csv,
            )
            task_status["progress"] = "完成"
        except Exception as e:
            task_status["progress"] = f"出错: {e}"
        finally:
            task_status["running"] = False

    thread = threading.Thread(target=run_in_background, daemon=True)
    thread.start()

    flash(f"开始发送 {len(recipients)} 封邮件（后台运行中），请稍候刷新页面", "info")
    return redirect(url_for("send_page"))


@app.route("/api/task-status")
def api_task_status():
    return jsonify(task_status)


# ---- 数据下载 ----

@app.route("/download/<filename>")
def download_file(filename):
    safe_names = {
        "references.csv", "authors_emails.csv",
        "failed_lookup.csv", "sent_log.csv", "manual_supplement.csv",
    }
    if filename not in safe_names:
        flash("无效的文件名", "danger")
        return redirect(url_for("index"))

    filepath = DATA_DIR / filename
    if not filepath.exists():
        flash(f"文件不存在: {filename}", "warning")
        return redirect(url_for("index"))

    return send_from_directory(str(DATA_DIR), filename, as_attachment=True)


# ==================== 启动 ====================

if __name__ == "__main__":
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    print("=" * 50)
    print("  《创新与创业教育》引用通知系统")
    print("  请在浏览器中打开: http://127.0.0.1:5000")
    print("=" * 50)
    app.run(debug=True, host="127.0.0.1", port=5000)
