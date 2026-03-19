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
YEAR_PATTERN = re.compile(r"[,，.]\s*((?:19|20)\d{2})\s*[,，(（]")
DOC_TYPE_PATTERN = re.compile(r"\[([JMCDPRSNEBOL/]+)\]")
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

        # 策略1：从脚注/尾注中提取参考文献
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
            "references": references,
            "ref_source": ref_source,
            "error": error,
        }

    except Exception as e:
        return {
            "source_file": os.path.basename(filepath),
            "article_title": "解析失败",
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
    article_title = _guess_title_from_paragraphs(lines[:10])
    references = _extract_refs_from_body(lines)

    return {
        "source_file": os.path.basename(filepath),
        "article_title": article_title,
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
    title_parts = []
    for p in paragraphs[:10]:
        if len(p) < 4:
            continue
        if any(kw in p for kw in ["摘要", "Abstract", "关键词", "基金", "收稿", "作者简介"]):
            break
        if 4 < len(p) < 100:
            title_parts.append(p)
            # 如果看起来像一个完整标题就停
            if len(p) > 10 and not p.endswith(("：", ":", "——")):
                break
    return " ".join(title_parts) if title_parts else (paragraphs[0] if paragraphs else "未知标题")


def _find_reference_section(paragraphs: list[str]) -> int | None:
    """查找参考文献章节的起始位置"""
    for i, p in enumerate(paragraphs):
        cleaned = p.replace(" ", "").replace("\u3000", "").strip()
        for marker in REFERENCE_SECTION_MARKERS:
            clean_marker = marker.replace(" ", "")
            if cleaned == clean_marker or cleaned.startswith(clean_marker):
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
        title = re.sub(r"\[[JMCDPRSNEBOL/]+\]", "", title).strip()
        result["title"] = title

        # 期刊名
        if len(parts) >= 3:
            journal = parts[2].strip()
            journal = re.sub(r"\[[JMCDPRSNEBOL/]+\]", "", journal).strip()
            journal = re.split(r"[,，]", journal)[0].strip()
            if journal and not re.match(r"^\d{4}", journal):
                result["journal"] = journal
    else:
        result["title"] = content


# ==================== 处理入口 ====================

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


def save_references_csv(results: list[dict], output_path: str):
    """将提取结果保存为 CSV"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    fieldnames = [
        "source_file", "article_title",
        "ref_number", "authors", "title", "year", "journal", "doi",
        "doc_type", "raw_text",
    ]

    rows = []
    for result in results:
        for ref in result["references"]:
            rows.append({
                "source_file": result["source_file"],
                "article_title": result["article_title"],
                **{k: ref.get(k, "") for k in fieldnames if k not in ("source_file", "article_title")},
            })

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"\n参考文献数据已保存到: {output_path}")
    print(f"共 {len(rows)} 条记录，来自 {len(results)} 篇文章")
    return rows
