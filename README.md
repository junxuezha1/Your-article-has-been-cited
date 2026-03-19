# 《创新与创业教育》引用通知系统

自动化学术引用通知系统，用于期刊编辑部向被引作者发送引用通知邮件。

## 功能特性

- 📄 **自动提取参考文献**：支持 Word (.doc/.docx) 和 PDF 格式
- 🔍 **智能邮箱检索**：通过 CrossRef 和 Semantic Scholar API 自动查找作者邮箱
- 📧 **批量邮件发送**：使用 HTML 模板批量发送个性化通知邮件
- 🌐 **Web 界面**：提供友好的网页操作界面
- 📊 **数据管理**：CSV 格式存储，支持人工补充和审核

## 系统架构

```
citation-notifier/
├── app.py                    # Flask Web 应用
├── main.py                   # 命令行入口
├── extract_references.py    # 参考文献提取模块
├── lookup_authors.py         # 作者邮箱检索模块
├── send_emails.py            # 邮件发送模块
├── config.yaml               # 配置文件（需自行创建）
├── requirements.txt          # Python 依赖
├── templates/                # HTML 模板
│   ├── notification.html     # 邮件通知模板
│   └── *.html                # Web 界面模板
└── data/                     # 数据目录（自动创建）
    ├── input/                # 输入文件目录
    ├── references.csv        # 提取的参考文献
    ├── authors_emails.csv    # 检索到的邮箱
    ├── failed_lookup.csv     # 检索失败的记录
    ├── manual_supplement.csv # 人工补充的邮箱
    └── sent_log.csv          # 发送日志
```

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

**注意**：Windows 系统需要安装 Microsoft Word 或 WPS Office 以支持 .doc/.docx 文件解析。

### 2. 配置系统

复制 `config.yaml` 并填写配置：

```yaml
smtp:
  server: "smtp.qq.com"
  port: 465
  use_ssl: true
  username: "your_email@qq.com"
  password: "your_auth_code"  # 邮箱授权码
  sender_name: "《创新与创业教育》编辑部"

crossref:
  mailto: "your_email@example.com"  # 用于 CrossRef Polite Pool

journal:
  name: "创新与创业教育"
  name_en: "Innovation and Entrepreneurship Education"
  website: "https://your-journal-website.com"
```

### 3. 启动 Web 界面（推荐）

```bash
python main.py web
```

访问 http://127.0.0.1:5000

### 4. 或使用命令行

```bash
# 提取参考文献
python main.py extract

# 检索作者邮箱
python main.py lookup

# 生成邮件预览
python main.py preview

# 发送邮件
python main.py send
```

## 使用流程

### 阶段一：提取参考文献

1. 将文章文件（.doc/.docx/.pdf）放入 `data/input/` 目录
2. 运行提取功能，系统自动识别参考文献章节
3. 结果保存到 `data/references.csv`

### 阶段二：检索作者邮箱

1. 系统通过以下 API 自动检索：
   - CrossRef（通过 DOI 或标题）
   - Semantic Scholar（通过标题）
2. 成功的记录保存到 `data/authors_emails.csv`
3. 失败的记录保存到 `data/failed_lookup.csv`

**人工补充**：
- 对于检索失败的记录，可在知网等数据库人工查找
- 填入 `data/manual_supplement.csv`（格式参考 `authors_emails.csv`）
- 运行合并功能：`python main.py merge`

### 阶段三：发送邮件

1. 预览邮件内容（Web 界面或 `python main.py preview`）
2. 确认无误后批量发送
3. 发送日志保存到 `data/sent_log.csv`

## 邮件模板

邮件模板位于 `templates/notification.html`，使用 Jinja2 语法，可自定义：

```html
<p>尊敬的 {{ author_name }} 老师：</p>
<p>您好！您的文章《{{ citation.title }}》在我刊被引用。</p>
```

可用变量：
- `author_name`：作者姓名
- `journal_name`：期刊名称
- `citations`：引用列表（包含 `article_title`, `title`, `year` 等）

## 技术栈

- **后端**：Python 3.8+, Flask
- **文档解析**：python-docx, pdfplumber, pywin32
- **学术 API**：CrossRef, Semantic Scholar
- **邮件发送**：smtplib (支持 SSL/TLS)
- **数据处理**：pandas

## 常见问题

### Q: Word 文件解析失败？
A: 确保已安装 Microsoft Word 或 WPS Office，系统使用 COM 自动化读取文档。

### Q: 邮箱检索成功率低？
A:
- 填写 CrossRef `mailto` 可提高请求优先级
- 申请 Semantic Scholar API Key 可提高配额
- 对于中文文献，建议人工补充

### Q: 邮件发送失败？
A:
- 检查 SMTP 配置是否正确
- 使用邮箱授权码而非登录密码
- 注意发送间隔，避免被判为垃圾邮件

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

如有问题，请通过 GitHub Issues 联系。
