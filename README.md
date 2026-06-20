# 《创新与创业教育》引用通知系统

自动化学术引用通知系统，用于期刊编辑部向被引作者发送引用通知邮件。

## 功能特性

- 📄 **自动提取参考文献**：支持 Word (`.doc/.docx`) 和 PDF 格式
- 🔍 **本地邮箱检索**：通过本地 Excel 作者库自动查找作者邮箱
- 📧 **批量邮件发送**：使用 HTML 模板批量发送个性化通知邮件
- 🌐 **Web 界面**：提供友好的网页操作界面
- 📊 **数据管理**：CSV 格式存储，支持人工补充和审核

## 系统架构

```text
citation-notifier/
├── app.py                      # Flask Web 应用
├── main.py                     # 命令行入口
├── extract_references.py       # 参考文献提取模块
├── lookup_authors.py           # 作者邮箱检索模块
├── db_lookup.py                # 本地作者库检索模块
├── send_emails.py              # 邮件发送模块
├── config.yaml                 # 配置文件（需自行创建）
├── requirements.txt            # Python 依赖
├── templates/                  # HTML 模板
│   ├── notification.html       # 邮件通知模板
│   └── *.html                  # Web 界面模板
├── 数据/                        # 本地作者库（.xls）
└── data/                       # 数据目录（自动创建）
    ├── input/                  # 输入文件目录
    ├── references.csv          # 提取的参考文献
    ├── references_no_author.csv# 暂未识别作者的文献
    ├── authors_emails.csv      # 检索到的邮箱
    ├── failed_lookup.csv       # 检索失败的记录
    ├── manual_supplement.csv   # 人工补充的邮箱
    └── sent_log.csv            # 发送日志
```

## 快速开始

### Windows 安装包（推荐）

普通用户可在 GitHub Release 页面下载 `CitationNotifier_Setup_v1.1.0.exe` 安装使用。安装包内置私有 Python 运行时，无需手动安装 Python 依赖。

安装后，程序数据与配置默认保存到：

```text
%LOCALAPPDATA%\CitationNotifier
```

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

**注意**：
- Windows 系统需要安装 Microsoft Word 或 WPS Office 以支持 `.doc/.docx` 文件解析。
- 本地作者库如果使用 `.xls` 文件，运行环境需要安装 `xlrd`。项目的 `requirements.txt` 已包含该依赖。

### 2. 配置系统

复制 `config.example.yaml` 为 `config.yaml` 并填写配置：

```yaml
smtp:
  server: "smtp.qq.com"
  port: 465
  use_ssl: true
  username: "your_email@qq.com"
  password: "your_auth_code"
  sender_name: "《创新与创业教育》编辑部"

app:
  secret_key: "change-this-secret-key"

journal:
  name: "创新与创业教育"
  name_en: "Innovation and Entrepreneurship Education"
  website: "https://your-journal-website.com"

paths:
  local_db_dir: "数据"
```

也可以通过环境变量 `CITATION_NOTIFIER_SECRET_KEY` 提供固定的 Web 密钥。

### 构建 Windows 安装器

开发者可在安装 Inno Setup 6 后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```

生成的安装包位于 `dist\installer\`。构建过程会下载 Python embeddable 运行时并安装 `requirements.txt` 中的依赖。

### 3. 启动 Web 界面（推荐）

```bash
python main.py web
```

访问 [http://127.0.0.1:5000](http://127.0.0.1:5000)

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

1. 将文章文件（`.doc/.docx/.pdf`）放入 `data/input/` 目录
2. 运行提取功能，系统自动识别参考文献章节
3. 结果保存到 `data/references.csv`

### 阶段二：检索作者邮箱

1. 系统从 `paths.local_db_dir` 指定的本地作者库目录加载 `.xls` 数据
2. 按参考文献中的作者姓名逐条匹配本地库中的邮箱
3. 成功的记录保存到 `data/authors_emails.csv`
4. 失败的记录保存到 `data/failed_lookup.csv`

**人工补充**：
- 对于检索失败的记录，可在知网等数据库人工查找
- 填入 `data/manual_supplement.csv`（格式参考 `authors_emails.csv`）
- 运行合并功能：`python main.py merge`

### 阶段三：发送邮件

1. 预览邮件内容（Web 界面或 `python main.py preview`）
2. 确认无误后批量发送
3. 发送日志保存到 `data/sent_log.csv`

## 邮件模板

邮件模板位于 `templates/notification.html`，使用 Jinja2 语法，可自定义。

可用变量：
- `author_name`：作者姓名
- `journal_name`：期刊名称
- `citations`：引用列表（包含 `cited_paper_title`、`citing_paper_title` 等）

## 技术栈

- **后端**：Python 3.8+, Flask
- **文档解析**：python-docx, pdfplumber, pywin32
- **本地作者库**：pandas + xlrd（读取 `.xls`）
- **邮件发送**：smtplib（支持 SSL/TLS）
- **数据处理**：pandas

## 常见问题

### Q: Word 文件解析失败？
A: 确保已安装 Microsoft Word 或 WPS Office，系统使用 COM 自动化读取文档。

### Q: 邮箱检索成功率低？
A:
- 检查 `paths.local_db_dir` 是否指向正确的本地作者库目录
- 确认作者库 `.xls` 文件完整，且已安装 `xlrd`
- 对于中文文献，建议人工补充

### Q: Web 界面里的登录态偶尔会失效？
A:
- 请在 `config.yaml` 的 `app.secret_key` 中设置固定密钥，或设置环境变量 `CITATION_NOTIFIER_SECRET_KEY`
- 如果未配置，系统会在每次启动时临时生成一个随机密钥

### Q: 邮件发送失败？
A:
- 检查 SMTP 配置是否正确
- 使用邮箱授权码而非登录密码
- 注意发送间隔，避免被判为垃圾邮件

## 许可证

MIT License
