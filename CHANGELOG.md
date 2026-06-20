# 修改与测试日志

本文件记录所有修改建议、测试结果和 Bug 修复，按时间顺序追加。

---

## 2026-04-22 — 初次审核与重构

### 需求变更

**用户要求**：
1. 删除 CrossRef / Semantic Scholar API 回退逻辑，完全依赖本地 Excel 数据库
2. 对三个阶段分别进行代码审核和测试

---

### 变更 1：lookup_authors.py — 删除外部 API

**文件**：`lookup_authors.py`

**删除内容**：
- `_query_crossref_by_doi()`
- `_query_crossref_by_title()`
- `_extract_crossref_author_info()`
- `_query_semantic_scholar()`
- `_get_ss_author_email()`
- `_titles_similar()`
- `import requests`、`import urllib.parse`

**保留内容**：本地数据库查询路径（`AuthorDatabase.lookup_by_name`）

**原因**：系统已有 69,253 条本地记录，外部 API 对中文作者命中率低，且引入网络依赖和延迟。

---

### 测试结果 — 第一阶段（参考文献提取）

**测试命令**：
```bash
python -c "from extract_references import _parse_single_reference, _looks_like_reference; print('import OK')"
```
**结果**：✅ 导入正常

**解析函数测试**：
```
[1] 张三,李四.创新创业教育研究综述[J].高等教育研究,2020,41(3):45-52.
  → authors='张三,李四', title='创新创业教育研究综述', year='2020', doi=''  ✅

[2] 王五.创业教育课程体系构建[M].北京:高等教育出版社,2019.
  → authors='王五', title='创业教育课程体系构建', year=''  ❌ 年份缺失

Smith J, Brown K. Innovation in education[J]. Journal of Education, 2021, 15(2): 100-110. DOI:10.1000/xyz123
  → authors='Smith J, Brown K', title='Innovation in education', year='2021', doi='10.1000/xyz123'  ✅
```

**Agent 评分**：3/5

---

### 测试结果 — 第二阶段（本地数据库检索）

**数据库加载**：
| 文件 | 行数 |
|---|---|
| persons (1).xls | 60 |
| persons (2).xls | 7,781 |
| persons (3).xls | 54,413 |
| 20260401115600.xls | 7,000 |
| **合计** | **69,253** |

**姓名查找测试**（精确匹配，5/5 命中）：
```
'谭桂林' → 'tanguilin322@sina.com'   ✅
'王昶'   → 'changw1000@163.com'      ✅
'李建华' → 'ljh5977@vip.sina.com'    ✅
'范武邱' → 'fanwuqiu@163.com'        ✅
'左高山' → 'mountaintso@126.com'     ✅
```

**真实数据测试**（references.csv，19 条记录）：
- `赵亮` → 命中 `ty33zl@yahoo.com.cn` ✅
- `陈寿灿`、`严毛新` → 未命中（不在库中）
- `UNESCO` → 未命中（正常）

**Agent 评分**：4/5

---

### 测试结果 — 第三阶段（邮件发送）

**模板渲染测试**：
```
Subject: 您的文章在《创新与创业教育》中被引用的通知  ✅
HTML 长度: 2732 字符  ✅
DOI='nan' 时 DOI 块正确隐藏  ✅
中文 Subject 自动 RFC 2047 编码  ✅
```

**prepare_email_data 测试**（data/authors_emails.csv）：
```
收件人数: 8，引用记录: 8
第一位: 赵亮 <ty33zl@yahoo.com.cn>，1 条引用  ✅
```

**Agent 评分**：3.5/5

---

### Bug 修复记录

#### BUG-001 ✅ 已修复
- **文件**：`extract_references.py` 第 25 行
- **严重度**：高
- **问题**：`YEAR_PATTERN` 要求年份后跟逗号或括号，书籍/学位论文格式（年份后跟句号）无法提取年份
- **修复**：正则改为 `r"[,，]\s*((?:19|20)\d{2})\s*[,，(（.]"`，末尾增加 `.` 匹配

#### BUG-002 ✅ 已修复
- **文件**：`extract_references.py` 第 229 行
- **严重度**：中
- **问题**：`cleaned.startswith(clean_marker)` 将"参考文献综述"误判为参考文献章节起始
- **修复**：`startswith` 后检查剩余字符，只允许标点（`：:。`）跟随，不允许汉字

#### BUG-003 ✅ 已修复
- **文件**：`extract_references.py` 第 416 行
- **严重度**：中
- **问题**：`os.makedirs(os.path.dirname(output_path))` 当路径无目录时传入空字符串，Windows 抛出 `FileNotFoundError`
- **修复**：`dir_name = os.path.dirname(output_path); if dir_name: os.makedirs(...)`

#### BUG-004 ✅ 已修复
- **文件**：`extract_references.py` 第 220 行
- **严重度**：中
- **问题**：`_guess_title_from_paragraphs` fallback 直接返回 `paragraphs[0]`，可能返回"作者简介：..."等元数据行
- **修复**：fallback 改为跳过包含元数据关键词的段落，取第一个干净段落

#### BUG-005 ✅ 已修复
- **文件**：`lookup_authors.py` 第 83 行
- **严重度**：高
- **问题**：`save_results` 中 `os.makedirs("")` 在纯文件名路径下崩溃
- **修复**：同 BUG-003 模式

#### BUG-006 ✅ 已修复
- **文件**：`lookup_authors.py` 第 71 行
- **严重度**：轻
- **问题**：`_parse_author_names` 未过滤"等"/"et al"等占位词，导致无意义数据库查询
- **修复**：增加 `skip_tokens = {"等", "et al", "et al.", "others"}` 过滤

#### BUG-007 ✅ 已修复
- **文件**：`lookup_authors.py` 第 19 行
- **严重度**：轻
- **问题**：数据库目录不存在时静默失败，无任何提示
- **修复**：增加 `print(f"[警告] 本地数据库目录不存在: {db_dir}...")`

#### BUG-008 ✅ 已修复
- **文件**：`send_emails.py` 第 35-41 行
- **严重度**：高
- **问题**：`matched_name='nan'` 是非空字符串，Python 视为 truthy，阻断 `corresponding_author` fallback，导致邮件称呼为空
- **修复**：引入 `_clean()` 函数，先将每个字段的 `'nan'` 清为空字符串再做 `or` 链

#### BUG-009 ✅ 已修复
- **文件**：`send_emails.py` 第 221 行
- **严重度**：高
- **问题**：`_append_sent_log` 中 `os.makedirs(os.path.dirname(sent_log_csv))` 空路径崩溃
- **修复**：同 BUG-003 模式

#### BUG-010 ✅ 已修复
- **文件**：`send_emails.py` 第 198-201 行
- **严重度**：中
- **问题**：`_make_log_key` 只用 `email + sorted(cited_paper_title)`，不含 `citing_paper_file`，跨期次同一引用被误判为已发送
- **修复**：key 改为 `email + sorted("cited_title@citing_file")`

#### BUG-011 ✅ 已修复
- **文件**：`send_emails.py` 第 163-164 行
- **严重度**：中
- **问题**：`finally: server.quit()` 若连接已断开会抛出新异常，抑制原始异常，难以调试
- **修复**：`finally: try: server.quit() except Exception: pass`

#### BUG-012 ✅ 已修复
- **文件**：`templates/notification.html` 第 57 行
- **严重度**：低
- **问题**：`author_name` 为空时显示" 老师/教授，您好！"（前置空格，称呼缺失）
- **修复**：`{% if author_name %}{{ author_name }} 老师/教授{% else %}尊敬的学者{% endif %}`

---

### 最终验证

```bash
python -c "
from extract_references import _parse_single_reference, _find_reference_section, save_references_csv
from lookup_authors import lookup_emails, save_results, _parse_author_names
from send_emails import prepare_email_data, render_email, _make_log_key

# _find_reference_section 修复验证
paras = ['参考文献综述', '参考文献', '[1] 张三.标题[J].期刊,2020.']
assert _find_reference_section(paras) == 1

# 年份提取修复验证（书籍格式）
r = _parse_single_reference('[2] 王五.创业教育课程体系构建[M].北京:高等教育出版社,2019.')
assert r['year'] == '2019'

# _parse_author_names 过滤'等'验证
assert '等' not in _parse_author_names('张三,李四,等')

# _make_log_key 包含 citing_paper_file 验证
r = {'email': 'a@b.com', 'citations': [{'cited_paper_title': 'T', 'citing_paper_file': 'f1.docx'}]}
assert 'f1.docx' in _make_log_key(r)

print('所有验证通过')
"
```
**结果**：所有验证通过

---

## 2026-04-22 — 三项功能改造（Agent Team）

### 需求

1. 三个板块联动：删除提取阶段的参考文献时，检索和发送阶段的对应数据也一并删除
2. 信息可编辑：检索结果和发送收件人均可手动修改邮箱、姓名
3. 选择性发送：默认全选，可手动取消勾选，发送按钮实时显示"发送已选 X 封"

### 改动文件

**`app.py`**（backend Agent）
- 新增 `_cascade_delete_by_titles(titles, config)` 辅助函数：从 `authors_emails.csv` 和 `failed_lookup.csv` 中删除 `title` 字段匹配的行
- 修改 `delete_ref_row`：删除前取出被删行的 title，删除后调用级联删除
- 修改 `delete_refs_batch`：批量取出所有被删行的 title 后级联删除
- 新增 `POST /lookup/edit-row`：按 idx 编辑 `authors_emails.csv` 中某行的 email、corresponding_author、matched_name
- 新增 `POST /send/edit-recipient`：按 original_email 匹配所有行，更新 email 和 matched_name
- 修改 `run_send`：接受 `selected_emails[]` 表单字段，有值则只发送选中邮箱，空则发送全部

**`templates/lookup.html`**（lookup-ui Agent）
- 页面描述改为"通过本地数据库检索被引作者的联系邮箱"
- "已找到邮箱"表格每行操作列加蓝色编辑按钮
- 新增编辑 Modal（editLookupModal），字段：邮箱、通讯作者、匹配姓名，POST 到 `/lookup/edit-row`
- 新增 `openEditLookup(idx, email, author, name)` JS 函数

**`templates/send_page.html`**（send-ui Agent）
- 所有收件人复选框默认 `checked`，全选框默认 `checked`
- 删除原批量删除工具栏，替换为已选数量提示
- 发送按钮改为"发送已选 X 封"，实时同步
- 新增 `updateSendCount()` 函数（替换原 `updateRecipientBar`）
- 新增 `prepareSendForm()` 函数：提交时将已勾选邮箱注入为 `selected_emails[]` hidden input
- 每个收件人项加蓝色编辑按钮
- 新增编辑 Modal（editRecipientModal），字段：邮箱地址、作者姓名，POST 到 `/send/edit-recipient`
- 新增 `openEditRecipient(email, authorName)` JS 函数

### 验证结果

```
OK /lookup/edit-row [POST]
OK /send/edit-recipient [POST]
OK run_send 支持 selected_emails[]
OK _cascade_delete_by_titles 存在
OK lookup.html: 描述更新 / Modal / JS / action
OK send_page.html: 选择字段 / prepareSend / updateCount / Modal / JS / action / 数量显示
```

---

---

## 2026-04-22 — 发送页四项功能改造

### 需求

1. 重复发送：点击发送键无论历史记录如何均可重新发送
2. 邮件加日期：签名区显示具体日期（年月日）
3. 手动添加收件人：无需上传 CSV，直接在页面输入邮箱和姓名
4. 邮件内容手工修改：每封邮件可单独编辑 HTML 源码，实时预览，保存后发送时使用修改版

### 改动文件

**`send_emails.py`**
- 删除 `_load_sent_log` 函数及 `send_emails` 中的去重跳过逻辑（`sent_set` 检查）
- `send_emails` 新增 `html_overrides: dict = None` 参数；发送前若该邮箱有覆盖内容则使用覆盖 HTML
- `render_email` 新增 `send_date=datetime.now().strftime("%Y年%m月%d日")` 传入模板

**`templates/notification.html`**
- 签名区由 `{{ year }} 年` 改为 `{{ send_date }}`

**`app.py`**
- `send_page()`：加载 `data/email_overrides.json`，渲染预览时应用覆盖内容
- `run_send()`：加载覆盖内容并通过 `html_overrides=` 传入 `send_emails()`
- 新增 `POST /send/add-recipient`：直接添加一条收件人记录到 `authors_emails.csv`
- 新增 `POST /send/save-email-override`：将编辑后的 HTML 写入 `data/email_overrides.json`

**`templates/send_page.html`**
- 发送记录区：默认高度 90px（约 3 行），可拖拽调整，最新记录在前，新增搜索框
- 新增"手动添加收件人"卡片（邮箱 + 姓名输入，POST 到 `/send/add-recipient`）
- 每个收件人项新增"编辑邮件内容"按钮（文档图标）
- 新增 `editEmailModal`：左侧 HTML 源码编辑区 + 右侧实时预览 iframe，POST 到 `/send/save-email-override`
- 新增 `openEditEmail(idx)` JS 函数；新增发送记录搜索逻辑

---

## 2026-04-23 — 发送日志 Excel 导出优化

### 需求

1. 发送日志包含"被引用文章"和"引用了的文章"列
2. 美化 Excel 输出（蓝色表头、成功绿色/失败红色行背景、冻结首行）
3. 导出按钮移到发送记录卡片下方

### 改动文件

**`send_emails.py`**
- `_append_sent_log` 改为按引用条目写行（每条引用一行）
- 字段由 `timestamp, email, author_name, citation_count, status, log_key` 改为 `timestamp, email, author_name, cited_paper_title, citing_paper_title, status`

**`app.py`**
- 顶部导入补充 `io`、`send_file`、`openpyxl.Workbook`、`openpyxl.styles.*`
- `get_data_stats()`：成功数改为 `drop_duplicates(subset=["timestamp","email"])` 去重计数，兼容新格式
- `send_page()`：sent_data 展示前按 `(timestamp, email)` 去重，避免多引用条目重复显示
- 新增 `GET /send/export-log`：读取 sent_log.csv，生成格式化 xlsx（深蓝表头、成功浅绿/失败浅红行、列宽自适应、冻结首行）

**`templates/send_page.html`**
- 发送记录卡片下方新增"导出发送记录 Excel"按钮（`/send/export-log`）

### 需求

用户要求：参考文献中的政策规定、新闻、网页、法律法规、案例等无个人作者的条目，应自动筛选出来，不进入邮箱检索阶段，但可由使用者手动选择是否加入检索。

### 改动文件

**`extract_references.py`**
- 新增常量 `NO_AUTHOR_DOC_TYPES`：`{EB/OL, N, S, Z, G, DB/OL, CP/CD}`（网页、报纸、标准、政府文件等）
- 新增常量 `_ORG_KEYWORDS`：机构名称特征词列表（国务院、教育部、全国、UNESCO 等）
- 新增函数 `_is_no_author_ref(ref)`：满足以下任一条件判定为无作者：
  1. `authors` 字段为空或 `'nan'`
  2. `doc_type` 属于 `NO_AUTHOR_DOC_TYPES`
  3. `authors` 包含机构名称特征词
- 修改 `save_references_csv(results, output_path, no_author_path=None)`：
  - 新增 `no_author_path` 参数
  - 自动将参考文献分流为两个 CSV 文件
  - 返回值改为 `(with_author, no_author)` 元组

**`config.yaml`**
- 新增路径配置：`references_no_author_csv: "data/references_no_author.csv"`

**`app.py`**
- `extract_page`：同时读取 `references.csv` 和 `references_no_author.csv`，传入模板
- `run_extract`：传入 `no_author_csv` 路径，接收分流结果，flash 消息显示两类数量
- 新增路由 `POST /extract/promote-no-author`：将单条无作者记录移入检索列表
- 新增路由 `POST /extract/promote-no-author-batch`：批量移入检索列表
- 新增路由 `POST /extract/delete-no-author`：从无作者列表删除单条
- `download_file`：白名单增加 `references_no_author.csv`

**`templates/extract.html`**
- 有作者表格标题改为"有作者参考文献（将进入邮箱检索阶段）"
- 新增无作者审核区（黄色警告样式）：
  - 显示条数 badge
  - 每行有"加入检索"（绿色箭头）和"删除"两个操作按钮
  - 支持全选 + 批量加入检索
  - 支持搜索筛选
  - 支持下载 `references_no_author.csv`

### 测试结果

```
✅ 空作者 → 无作者
✅ nan作者 → 无作者
✅ 网页类型(EB/OL) → 无作者
✅ 机构名称(国务院) → 无作者
✅ 正常期刊(张三,[J]) → 有作者
✅ 正常书籍(张三,[M]) → 有作者
✅ 政策文件(教育部,[S]) → 无作者
✅ app 导入正常，所有新路由注册成功
```

---

## 2026-04-22 — SMTP 配置与邮件格式优化

### 问题排查

**BUG-013 ✅ 已修复**
- **文件**：`send_emails.py` 第 11-12 行
- **严重度**：高
- **问题**：`From` 头直接拼接中文发件人名称（如"《创新与创业教育》编辑部"），未按 RFC2047 编码，QQ SMTP 服务器返回 `550 The "From" header is missing or invalid`
- **修复**：引入 `email.header.Header` 和 `email.utils.formataddr`，改为 `formataddr((str(Header(sender_name, 'utf-8')), sender_email))`

**BUG-014 ✅ 已修复**
- **文件**：`extract_references.py` `_guess_title_from_paragraphs`
- **严重度**：中
- **问题**：文档首段为 DOI 行（如 `DOI: 10.11817/...`）时被误识别为文章标题，导致邮件中"引用该文献的本刊文章"显示 DOI 而非标题
- **修复**：新增 `SKIP_PATTERNS` 正则，跳过以 `DOI/doi/https/www/issn` 开头的段落

### 新增功能

**新增 `article_authors` 字段**
- **文件**：`extract_references.py`、`send_emails.py`、`templates/notification.html`
- **说明**：新增 `_guess_authors_from_paragraphs()` 函数，从文档标题后的短段落中提取引用文章的作者；写入 CSV 的 `article_authors` 列；邮件模板中"引用该文献的本刊文章"区块展示标题 + 作者

### 邮件模板重构（`templates/notification.html`）

- 标题区居中显示
- 落款（编辑部 + 年份）改为右对齐
- 引用块分上下两部分，中间加分割线，去掉 DOI 显示
- 被引文章和引用文章均展示标题 + 作者（作者字段为空时自动隐藏）
- 整体视觉优化：圆角、背景色、字号层次

---

## 2026-04-23 — 邮件模板称呼与附件说明调整

### 改动文件

**`templates/notification.html`**
- 称呼由"老师/教授"改为"教授"
- 在"此致"前新增一段：「我们在附件中给您发送了原文的word版本，如您在后续大作中引用，请以知网PDF版本为准。」

---

## 2026-04-25 — 文档对齐、依赖补齐与发送日志兼容修复

### 需求

1. 修正文档、配置样例和页面文案中已落后的说明，使其与当前“本地 Excel 作者库检索”实现一致
2. 解决 `.xls` 本地作者库读取所需依赖缺失的问题
3. 不处理第 3 项
4. 修复发送日志新旧字段不一致导致的兼容问题
5. 修复 Flask `secret_key` 硬编码问题，改为可配置

### 改动文件

**`README.md`**
- 将“CrossRef / Semantic Scholar API 检索”整体改为“本地 Excel 作者库检索”
- 更新系统架构说明，补充 `db_lookup.py`、`数据/`、`references_no_author.csv`
- 快速开始中补充 `xlrd` 依赖说明
- 配置示例中补充 `app.secret_key` 与 `paths.local_db_dir`
- FAQ 中改写邮箱检索说明，并补充 Web 密钥说明

**`config.example.yaml`**
- 删除过时的 `crossref` / `semantic_scholar` 示例配置
- 新增 `app.secret_key`
- 在 `paths` 中补充 `references_no_author_csv` 与 `local_db_dir`

**`requirements.txt`**
- 删除未再使用的 `requests`
- 新增 `xlrd==2.0.1`，用于读取本地 `.xls` 作者库

**`db_lookup.py`**
- 模块头注释改为“本地 Excel 作者库检索”，去除外部 API 相关描述

**`templates/config_page.html`**
- 页面说明由“SMTP + 学术 API 配置”改为“SMTP + 本地作者库配置”
- 删除 CrossRef 配置卡片
- 改为配置 `local_db_dir`

**`templates/index.html`**
- 首页第 2 步文案改为“通过本地作者库自动检索被引作者的联系邮箱”

**`templates/lookup.html`**
- 检索说明改为“在本地作者库中检索通讯作者邮箱”

**`send_emails.py`**
- 新增 `SENT_LOG_FIELDS` / `LEGACY_SENT_LOG_FIELDS`
- 新增 `normalize_sent_log_df()`：统一新旧发送日志字段
- 新增 `load_sent_log()`：兼容读取旧版 `sent_log.csv`
- 新增 `_ensure_sent_log_schema()`：发送前若检测到旧格式日志，先迁移为新字段后再追加写入
- 保留旧日志中的 `timestamp / email / author_name / status`，并尽量从 `log_key` 还原 `cited_paper_title`

**`app.py`**
- 顶部导入改为使用 `load_sent_log`
- `get_data_stats()`、`send_page()`、`export_sent_log()` 全部改为使用兼容日志读取
- 新增 `resolve_secret_key()`：优先读取环境变量 `CITATION_NOTIFIER_SECRET_KEY`，其次读取 `config.yaml` 的 `app.secret_key`，否则临时生成随机密钥
- 配置页保存逻辑改为保存 `paths.local_db_dir`
- 去除保存 `crossref_mailto` 的逻辑

### 验证结果

```
OK 已将修复后的文件同步回原目录 E:\桌面\成功\Your-article-has-been-cited
OK app.py 已包含 resolve_secret_key / load_sent_log / local_db_dir
OK send_emails.py 已包含新旧发送日志兼容逻辑
OK README / config.example / 配置页 / 首页 / 检索页文案已切换为本地作者库方案
OK requirements.txt 已补充 xlrd==2.0.1
OK 使用旧版 sent_log.csv 样本验证：可兼容读取，并可在迁移后继续追加新记录
OK 修改副本通过 py_compile 基础语法检查
```

### 说明

- 本次未修改真实 `config.yaml`，其中保留的历史 `crossref` / `semantic_scholar` 字段不会影响当前代码运行
- `CHANGELOG.md` 中仍保留历史 CrossRef / Semantic Scholar 记录，作为过往迭代说明
