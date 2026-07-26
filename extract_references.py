"""
第一阶段：从 Word (.doc/.docx) / PDF 文件中提取参考文献
使用 pywin32 COM 调用 Word/WPS 读取文档，兼容 .doc 和 .docx 格式。
"""

import os
import re
import csv
from pathlib import Path

import pdfplumber


# 匹配参考文献章节标记
REFERENCE_SECTION_MARKERS = [
    "参考文献", "参 考 文 献", "References", "REFERENCES",
    "参考文献：", "参考文献:", "[参考文献]",
]

# 匹配 [1] [2] 等编号，也匹配无编号的文献条目
REF_NUM_PATTERN = re.compile(r"^\[(\d+)\]\s*(.+)")

# 提取结构化字段的正则
DOI_PATTERN = re.compile(r"(?:DOI|doi)[:\s]*?(10\.\d{4,}/[^\s,;]+)")
# 匹配期刊格式年份（年份后跟逗号/括号），以及书籍/学位论文格式（年份后跟句号）
YEAR_PATTERN = re.compile(r"[,，]\s*((?:19|20)\d{2})\s*[,，(（.]")
DOC_TYPE_PATTERN = re.compile(r"\[([A-Z]+(?:/[A-Z]+)?)\]")
URL_PATTERN = re.compile(r"https?://[^\s,，。]+")


def extract_from_word(filepath: str) -> dict:
    """
    用 COM 自动化（Word 或 WPS）读取 .doc/.docx 文件，
    提取文章标题、正文段落和脚注/尾注中的参考文献。
    """
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        abs_path = os.path.abspath(filepath)
        doc = word.Documents.Open(abs_path, ReadOnly=True)

        # 提取段落文本
        paragraphs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            text = doc.Paragraphs(i).Range.Text.strip()
            if text:
                paragraphs.append(text)

        # 提取脚注
        footnotes = []
        for i in range(1, doc.Footnotes.Count + 1):
            text = doc.Footnotes(i).Range.Text.strip()
            if text:
                footnotes.append(text)

        # 提取尾注
        endnotes = []
        for i in range(1, doc.Endnotes.Count + 1):
            text = doc.Endnotes(i).Range.Text.strip()
            if text:
                endnotes.append(text)

        article_title = _guess_title_from_paragraphs(paragraphs)
        article_authors = _guess_authors_from_paragraphs(paragraphs)
        notes = footnotes + endnotes
        refs_from_notes = _extract_refs_from_notes(notes)

        # 策略2：从正文末尾的"参考文献"章节提取
        refs_from_body = _extract_refs_from_body(paragraphs)

        # 取数量更多的那个结果（避免误判）
        if len(refs_from_notes) >= len(refs_from_body):
            references = refs_from_notes
            ref_source = "脚注/尾注"
        else:
            references = refs_from_body
            ref_source = "正文参考文献章节"

        error = None
        if not references:
            error = "未找到参考文献（已检查正文、脚注和尾注）"

        return {
            "source_file": os.path.basename(filepath),
            "article_title": article_title,
            "article_authors": article_authors,
            "references": references,
            "ref_source": ref_source,
            "error": error,
        }

    except Exception as e:
        return {
            "source_file": os.path.basename(filepath),
            "article_title": "解析失败",
            "article_authors": "",
            "references": [],
            "ref_source": "",
            "error": str(e),
        }
    finally:
        try:
            if doc:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def extract_from_pdf(filepath: str) -> dict:
    """从 PDF 文件中提取参考文献"""
    full_text = ""
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    lines = [line.strip() for line in full_text.split("\n") if line.strip()]
    article_title = _guess_title_from_paragraphs(lines[:15])
    article_authors = _guess_authors_from_paragraphs(lines[:20])
    references = _extract_refs_from_body(lines)

    return {
        "source_file": os.path.basename(filepath),
        "article_title": article_title,
        "article_authors": article_authors,
        "references": references,
        "ref_source": "PDF正文",
        "error": None if references else "未找到参考文献章节",
    }


# ==================== 参考文献提取策略 ====================

def _extract_refs_from_notes(notes: list[str]) -> list[dict]:
    """从脚注/尾注列表中提取参考文献（每条脚注可能就是一条参考文献）"""
    refs = []
    for i, text in enumerate(notes):
        # 脚注中包含学术引用特征：有 [J] [M] 等标记，或年份+期刊格式
        if _looks_like_reference(text):
            parsed = _parse_single_reference(text, ref_number=i + 1)
            refs.append(parsed)
    return refs


def _extract_refs_from_body(paragraphs: list[str]) -> list[dict]:
    """从正文的"参考文献"章节提取"""
    ref_start_idx = _find_reference_section(paragraphs)
    if ref_start_idx is None:
        return []

    # 收集参考文献章节之后的条目（在英文摘要/编辑信息之前截止）
    ref_lines = []
    for line in paragraphs[ref_start_idx + 1:]:
        # 遇到英文标题、编辑信息等则停止
        if _is_end_of_references(line):
            break
        ref_lines.append(line)

    raw_refs = _collect_reference_entries(ref_lines)
    return [_parse_single_reference(r) for r in raw_refs]


def _looks_like_reference(text: str) -> bool:
    """判断一段文本是否像学术参考文献"""
    # 包含文献类型标记 [J] [M] [C] [D] [EB/OL] 等
    if DOC_TYPE_PATTERN.search(text):
        return True
    # 包含年份 + 逗号/句号的组合
    if re.search(r"(19|20)\d{2}", text) and len(text) > 20:
        # 并且包含句号（分隔作者和标题）
        if re.search(r"[.．。]", text):
            return True
    return False


def _is_end_of_references(line: str) -> bool:
    """判断是否到了参考文献章节的结束位置"""
    end_markers = [
        "Abstract", "Key words", "Keywords", "[编辑", "（编辑",
        "收稿日期", "基金项目", "作者简介",
    ]
    for marker in end_markers:
        if line.strip().startswith(marker):
            return True
    # 全英文标题行（可能是英文摘要的标题）
    if re.match(r"^[A-Z][a-z]+([\s\-][a-zA-Z]+){2,}", line.strip()):
        if not DOC_TYPE_PATTERN.search(line):
            return True
    return False


# ==================== 工具函数 ====================

def _guess_title_from_paragraphs(paragraphs: list[str]) -> str:
    """从前几段猜测文章标题"""
    META_KEYWORDS = ["摘要", "Abstract", "关键词", "基金", "收稿", "作者简介"]
    SKIP_PATTERNS = re.compile(r"^(DOI|doi|https?://|www\.|issn|ISSN)", re.IGNORECASE)
    title_parts = []
    candidates = paragraphs[:15]
    for idx, p in enumerate(candidates):
        if len(p) < 4:
            continue
        if SKIP_PATTERNS.match(p):
            continue
        if any(kw in p for kw in META_KEYWORDS):
            break
        if 4 < len(p) < 100:
            title_parts.append(p)
            next_p = candidates[idx + 1].strip() if idx + 1 < len(candidates) else ""
            if len(p) > 10 and not p.endswith(("：", ":", "——")) and not _looks_like_title_continuation(next_p):
                break
    return _join_title_parts(title_parts) if title_parts else next(
        (p for p in paragraphs if p and not any(kw in p for kw in META_KEYWORDS) and not SKIP_PATTERNS.match(p)),
        "未知标题",
    )


def _join_title_parts(parts: list[str]) -> str:
    """合并标题与副标题，避免破折号副标题前出现多余空格。"""
    title = ""
    for part in parts:
        part = part.strip()
        if not title:
            title = part
        elif _looks_like_title_continuation(part):
            title += part
        else:
            title += " " + part
    return title


def _looks_like_title_continuation(text: str) -> bool:
    stripped = text.strip()
    return stripped.startswith(("——", "—", "--", "副标题"))


def _looks_like_affiliation_line(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if stripped.startswith(("(", "（")) and stripped.endswith((")", "）")):
        return True
    affiliation_keywords = [
        "大学", "学院", "研究院", "研究所", "中心", "实验室", "系",
        "School", "University", "Institute", "College",
    ]
    return any(kw in stripped for kw in affiliation_keywords) and re.search(r"\d{5,6}|北京|上海|南京|长沙|广州|武汉|成都", stripped)


def _looks_like_author_line(text: str) -> bool:
    stripped = text.strip()
    if not 1 < len(stripped) < 60:
        return False
    if _looks_like_title_continuation(stripped) or _looks_like_affiliation_line(stripped):
        return False
    if re.search(r"[。；;：:?？!！]", stripped):
        return False
    if any(kw in stripped for kw in ("摘要", "关键词", "基金", "收稿", "作者简介")):
        return False
    # 去掉常见脚注/通讯作者标记后，作者行应主要由姓名分隔符组成。
    cleaned = re.sub(r"[\d¹²³⁴⁵⁶⁷⁸⁹⁰*＊†‡]", "", stripped)
    cleaned = re.sub(r"\s+", "", cleaned)
    return bool(re.fullmatch(r"[\u4e00-\u9fffA-Za-z·.,，、\-]+", cleaned))


def _clean_author_line(text: str) -> str:
    cleaned = re.sub(r"[\d¹²³⁴⁵⁶⁷⁸⁹⁰*＊†‡]", "", text.strip())
    cleaned = re.sub(r"\s+", "", cleaned)
    return cleaned.strip("，,、")


def _guess_authors_from_paragraphs(paragraphs: list[str]) -> str:
    """从前几段猜测文章作者（标题之后、摘要之前的短段落）"""
    META_KEYWORDS = ["摘要", "Abstract", "关键词", "基金", "收稿"]
    SKIP_PATTERNS = re.compile(r"^(DOI|doi|https?://|www\.|issn|ISSN)", re.IGNORECASE)
    # 找到标题后的段落，作者通常是短段落（< 50字），不含标点句子
    found_title = False
    title_needs_continuation = False
    for p in paragraphs[:20]:
        if not p or SKIP_PATTERNS.match(p):
            continue
        if any(kw in p for kw in META_KEYWORDS):
            break
        if not found_title:
            if len(p) < 4:
                continue
            found_title = True
            title_needs_continuation = p.endswith(("：", ":", "——"))
            continue
        if title_needs_continuation:
            title_needs_continuation = p.endswith(("：", ":", "——"))
            continue
        # 作者行特征：较短、含中文姓名或逗号分隔
        if _looks_like_author_line(p):
            return _clean_author_line(p)
    return ""


def _find_reference_section(paragraphs: list[str]) -> int | None:
    """查找参考文献章节的起始位置"""
    # 允许 marker 后跟标点，但不允许跟汉字（避免"参考文献综述"误判）
    trailing_ok = re.compile(r"^[：:。\s]*$")
    for i, p in enumerate(paragraphs):
        cleaned = p.replace(" ", "").replace("　", "").strip()
        for marker in REFERENCE_SECTION_MARKERS:
            clean_marker = marker.replace(" ", "")
            if cleaned == clean_marker:
                return i
            if cleaned.startswith(clean_marker):
                remainder = cleaned[len(clean_marker):]
                if trailing_ok.match(remainder):
                    return i
    return None

def _collect_reference_entries(lines: list[str]) -> list[str]:
    """收集参考文献条目，处理有编号和无编号两种情况"""
    entries = []
    current = ""

    # 先判断是否有编号格式
    has_numbering = any(REF_NUM_PATTERN.match(line.strip()) for line in lines if line.strip())

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if has_numbering:
            # 有编号模式：[1] [2] 开头的是新条目，其余合并到上一条
            if REF_NUM_PATTERN.match(line):
                if current:
                    entries.append(current.strip())
                current = line
            elif current:
                current += " " + line
        else:
            # 无编号模式：每一行如果看起来像参考文献就是独立的一条
            if _looks_like_reference(line):
                if current:
                    entries.append(current.strip())
                current = line
            elif current:
                # 可能是上一条的续行
                current += " " + line

    if current:
        entries.append(current.strip())

    return entries


def _parse_single_reference(raw_text: str, ref_number: int | None = None) -> dict:
    """解析单条参考文献文本"""
    result = {
        "raw_text": raw_text,
        "ref_number": ref_number,
        "authors": "",
        "title": "",
        "year": "",
        "journal": "",
        "doi": "",
        "doc_type": "",
    }

    # 提取编号（如果文本中有 [数字] 开头）
    num_match = REF_NUM_PATTERN.match(raw_text)
    if num_match:
        result["ref_number"] = int(num_match.group(1))
        content = num_match.group(2)
    else:
        content = raw_text

    # 提取 DOI
    doi_match = DOI_PATTERN.search(content)
    if doi_match:
        result["doi"] = doi_match.group(1).rstrip(".")

    # 提取文献类型 [J] [M] [EB/OL] 等
    type_match = DOC_TYPE_PATTERN.search(content)
    if type_match:
        result["doc_type"] = type_match.group(1)

    # 提取年份
    year_match = YEAR_PATTERN.search(content)
    if year_match:
        result["year"] = year_match.group(1)
    else:
        # 备用：匹配括号中的年份
        alt_year = re.search(r"[(（]\s*((?:19|20)\d{2})", content)
        if alt_year:
            result["year"] = alt_year.group(1)

    # 提取作者和标题
    _extract_authors_and_title(content, result)

    return result


def _extract_authors_and_title(content: str, result: dict):
    """从参考文献内容中分离作者和标题"""
    # 中文格式: 作者1, 作者2. 标题[J]. 期刊名, 年, 卷(期): 页.
    # 英文格式: Author1, Author2. Title[J]. Journal, Year, Vol(No): Pages.
    # 先去掉 URL 部分避免干扰
    clean = URL_PATTERN.sub("", content)
    # 按句号分割（中文句号和英文句号）
    parts = re.split(r"[.．。]\s*", clean, maxsplit=3)

    if len(parts) >= 2:
        result["authors"] = parts[0].strip().rstrip(",，、")

        # 标题部分：去掉 [J] 等标记
        title = parts[1].strip()
        title = DOC_TYPE_PATTERN.sub("", title).strip()
        result["title"] = title

        # 期刊名
        if len(parts) >= 3:
            journal = parts[2].strip()
            journal = DOC_TYPE_PATTERN.sub("", journal).strip()
            journal = re.split(r"[,，]", journal)[0].strip()
            if journal and not re.match(r"^\d{4}", journal):
                result["journal"] = journal
    else:
        result["title"] = content


# ==================== 处理入口 ====================

# 匹配中文字符（CJK 统一汉字）
_CJK_RE = re.compile(r'[一-鿿㐀-䶿]')

# 只有中文期刊默认进入邮箱检索，其余类型进入待审核区。
JOURNAL_DOC_TYPES = {"J", "J/OL"}
BOOK_DOC_TYPES = {"M", "M/OL"}
POLICY_DOC_TYPES = {"S", "Z", "G"}
WEB_NEWS_DOC_TYPES = {"EB/OL", "N", "DB/OL", "CP/CD"}
ACADEMIC_NON_JOURNAL_DOC_TYPES = {"C", "D", "R", "P"}

# 机构名称特征词（出现则视为无个人作者）
_ORG_KEYWORDS = [
    "国务院", "教育部", "人大", "全国", "中共", "中央", "政府", "委员会",
    "人民法院", "最高法", "法院", "检察院", "公安部", "财政部", "发改委",
    "国家", "省", "市政府", "人民政府", "办公厅", "新华社", "人民日报",
    "中国", "联合国", "UNESCO", "WHO", "OECD",
]

_POLICY_KEYWORDS = [
    "条例", "办法", "意见", "通知", "决定", "规划", "方案", "纲要",
    "法律", "法规", "政策", "标准", "规定", "报告", "公报", "白皮书",
]

_BOOK_PUBLISHER_KEYWORDS = [
    "出版社", "出版集团", "出版公司", "出版传媒", "书局", "Press",
    "Publishing", "Publisher",
]

_BOOK_PLACE_PATTERN = re.compile(
    r"(北京|上海|天津|重庆|南京|武汉|广州|长沙|杭州|成都|西安|济南|郑州|"
    r"合肥|福州|南昌|长春|沈阳|哈尔滨|石家庄|太原|兰州|昆明|贵阳|南宁|"
    r"呼和浩特|乌鲁木齐|海口|银川|拉萨|香港|台北)\s*[:：]"
)

_JOURNAL_VOLUME_PATTERN = re.compile(
    r"(?:19|20)\d{2}\s*[,，]\s*\d+\s*(?:\(\d+\)|（\d+）|[,，:：])"
)


def _normalize_doc_type(doc_type: str) -> str:
    return str(doc_type or "").strip().upper().replace(" ", "")


def _has_org_author(authors: str) -> bool:
    return any(kw in authors for kw in _ORG_KEYWORDS)


def _looks_like_book_ref(ref: dict) -> bool:
    """识别未规范标注为 [M] 但出版形态明显是图书的参考文献。"""
    doc_type = _normalize_doc_type(ref.get("doc_type", ""))
    if doc_type in BOOK_DOC_TYPES:
        return True

    raw = str(ref.get("raw_text", "")).strip()
    journal = str(ref.get("journal", "")).strip()
    text = f"{raw} {journal}"

    if re.search(r"\[M(?:/[A-Z]+)?\]", raw, re.IGNORECASE):
        return True
    if any(kw in text for kw in _BOOK_PUBLISHER_KEYWORDS):
        return True
    if _BOOK_PLACE_PATTERN.search(text) and re.search(r"[:：]\s*[^,，。]*(出版|Press|Publishing)", text, re.IGNORECASE):
        return True
    if re.search(r"(第\s*\d+\s*版|主编|译著|译\.)", raw):
        return True

    return False


def _looks_like_policy_ref(ref: dict) -> bool:
    doc_type = _normalize_doc_type(ref.get("doc_type", ""))
    if doc_type in POLICY_DOC_TYPES:
        return True

    authors = str(ref.get("authors", "")).strip()
    title = str(ref.get("title", "")).strip()
    raw = str(ref.get("raw_text", "")).strip()
    text = f"{authors} {title} {raw}"

    if _has_org_author(authors) and any(kw in text for kw in _POLICY_KEYWORDS):
        return True
    if re.search(r"(国务院|教育部|中共中央|全国人大|人民政府|办公厅).*(通知|意见|办法|条例|规划|决定|方案)", text):
        return True

    return False


def _looks_like_journal_ref(ref: dict) -> bool:
    """判断是否可作为中文期刊进入自动邮箱检索。"""
    doc_type = _normalize_doc_type(ref.get("doc_type", ""))
    if doc_type in JOURNAL_DOC_TYPES:
        return not _looks_like_book_ref(ref)

    if doc_type:
        return False

    raw = str(ref.get("raw_text", "")).strip()
    journal = str(ref.get("journal", "")).strip()

    if _looks_like_book_ref(ref) or _looks_like_policy_ref(ref):
        return False

    # 无 [J] 标识时，用期刊常见结构兜底：刊名 + 年份 + 卷(期)/页码。
    if journal and _JOURNAL_VOLUME_PATTERN.search(raw):
        return True
    if re.search(r"(学报|期刊|杂志|研究|论坛|教育|科学|社会科学|Journal|Review|Quarterly)", journal, re.IGNORECASE):
        if re.search(r"(?:19|20)\d{2}", raw):
            return True

    return False


def _classify_reference(ref: dict) -> str:
    """返回用于分流和展示的参考文献类别。"""
    if _is_foreign_ref(ref):
        return "外文"

    doc_type = _normalize_doc_type(ref.get("doc_type", ""))

    if _looks_like_policy_ref(ref):
        return "政策/法规"
    if _looks_like_book_ref(ref):
        return "书籍"
    if doc_type in WEB_NEWS_DOC_TYPES:
        return "网页/新闻"
    if doc_type in ACADEMIC_NON_JOURNAL_DOC_TYPES:
        return "专利/会议/学位/报告"
    if _looks_like_journal_ref(ref):
        return "期刊"

    authors = str(ref.get("authors", "")).strip()
    if not authors or authors == "nan" or _has_org_author(authors):
        return "无个人作者"

    return "其他"


def _is_no_author_ref(ref: dict) -> bool:
    """
    判断一条参考文献是否应排除出自动邮箱检索。
    目前仅中文期刊默认进入检索，其他类型进入待审核区。
    """
    return _classify_reference(ref) != "期刊"


def _is_foreign_ref(ref: dict) -> bool:
    """
    判断是否为外文文献：作者和标题均不含中文字符，且有实际内容。
    外文文献不进入自动检索，归入待审核区供人工选择。
    """
    authors = str(ref.get("authors", "")).strip()
    title = str(ref.get("title", "")).strip()
    raw = str(ref.get("raw_text", "")).strip()
    # 有中文字符 → 中文文献
    if _CJK_RE.search(authors) or _CJK_RE.search(title):
        return False
    # 无中文且有实质内容 → 外文文献
    return bool(authors or title or raw)


WORD_EXTENSIONS = {".doc", ".docx"}

def process_input_directory(input_dir: str) -> list[dict]:
    """处理输入目录中的所有 Word 和 PDF 文件"""
    results = []
    input_path = Path(input_dir)

    if not input_path.exists():
        print(f"错误：输入目录不存在 — {input_dir}")
        return results

    files = sorted(input_path.iterdir())
    word_files = [f for f in files if f.suffix.lower() in WORD_EXTENSIONS and not f.name.startswith("~")]
    pdf_files = [f for f in files if f.suffix.lower() == ".pdf"]

    print(f"发现 {len(word_files)} 个 Word 文件, {len(pdf_files)} 个 PDF 文件")

    for f in word_files:
        print(f"  正在处理: {f.name}")
        try:
            result = extract_from_word(str(f))
            results.append(result)
            ref_count = len(result["references"])
            source = result.get("ref_source", "")
            print(f"    ✓ 提取到 {ref_count} 条参考文献（来源: {source}）")
            if result["error"]:
                print(f"    ⚠ {result['error']}")
        except Exception as e:
            print(f"    ✗ 处理失败: {e}")
            results.append({
                "source_file": f.name,
                "article_title": "解析失败",
                "references": [],
                "error": str(e),
            })

    for f in pdf_files:
        print(f"  正在处理: {f.name}")
        try:
            result = extract_from_pdf(str(f))
            results.append(result)
            ref_count = len(result["references"])
            print(f"    ✓ 提取到 {ref_count} 条参考文献")
            if result["error"]:
                print(f"    ⚠ {result['error']}")
        except Exception as e:
            print(f"    ✗ 处理失败: {e}")
            results.append({
                "source_file": f.name,
                "article_title": "解析失败",
                "references": [],
                "error": str(e),
            })

    return results


def split_reference_rows(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    """Split normalized reference rows into papers and non-notification references."""
    papers = []
    other = []
    for row in rows:
        ref_category = _classify_reference(row)
        row["ref_lang"] = "外文" if _is_foreign_ref(row) else "中文"
        row["ref_category"] = ref_category
        if ref_category == "期刊":
            papers.append(row)
        else:
            other.append(row)
    return papers, other


def save_references_csv(results: list[dict], output_path: str, no_author_path: str = None):
    """
    将提取结果保存为 CSV，同时按有无个人作者分流：
    - output_path：有作者的参考文献（进入检索阶段）
    - no_author_path：无作者/机构类文献（供人工审核后决定是否加入检索）
    """
    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)

    fieldnames = [
        "source_file", "article_title", "article_authors",
        "ref_number", "authors", "title", "year", "journal", "doi",
        "doc_type", "ref_lang", "ref_category", "raw_text",
    ]

    rows = []

    for result in results:
        for ref in result["references"]:
            row = {
                "source_file": result["source_file"],
                "article_title": result["article_title"],
                "article_authors": result.get("article_authors", ""),
                **{k: ref.get(k, "") for k in fieldnames
                   if k not in ("source_file", "article_title", "article_authors", "ref_lang", "ref_category")},
            }
            rows.append(row)

    with_author, no_author = split_reference_rows(rows)

    def _write(path, rows):
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

    _write(output_path, with_author)
    print(f"\n有作者参考文献已保存: {output_path}（{len(with_author)} 条）")

    if no_author_path:
        no_author_dir = os.path.dirname(no_author_path)
        if no_author_dir:
            os.makedirs(no_author_dir, exist_ok=True)
        _write(no_author_path, no_author)
        print(f"无作者参考文献已保存: {no_author_path}（{len(no_author)} 条，待人工审核）")

    print(f"合计 {len(with_author) + len(no_author)} 条，来自 {len(results)} 篇文章")
    return with_author, no_author

