"""
第二阶段：通过学术数据库 API 检索被引作者的联系邮箱
"""

import csv
import os
import time
import urllib.parse

import pandas as pd
import requests


def lookup_emails(references_csv: str, config: dict) -> pd.DataFrame:
    """
    读取参考文献 CSV，逐条通过 API 检索作者邮箱。
    返回包含邮箱信息的 DataFrame。
    """
    df = pd.read_csv(references_csv, encoding="utf-8-sig")
    print(f"共加载 {len(df)} 条参考文献记录")

    # 新增列
    df["corresponding_author"] = ""
    df["email"] = ""
    df["lookup_source"] = ""
    df["lookup_status"] = ""

    crossref_cfg = config.get("crossref", {})
    ss_cfg = config.get("semantic_scholar", {})

    for idx, row in df.iterrows():
        title = str(row.get("title", "")).strip()
        doi = str(row.get("doi", "")).strip()
        authors = str(row.get("authors", "")).strip()

        if not title and not doi:
            df.at[idx, "lookup_status"] = "跳过：标题和DOI均为空"
            continue

        print(f"  [{idx + 1}/{len(df)}] 检索: {title[:50]}...")

        # 策略1: 如果有 DOI，先用 DOI 查 CrossRef
        if doi:
            result = _query_crossref_by_doi(doi, crossref_cfg)
            if result and result.get("email"):
                _fill_result(df, idx, result, "CrossRef-DOI")
                continue

        # 策略2: 用标题查 CrossRef
        if title:
            result = _query_crossref_by_title(title, crossref_cfg)
            if result and result.get("email"):
                _fill_result(df, idx, result, "CrossRef-Title")
                continue
            # 即使没有邮箱，如果匹配到了DOI也记录
            if result and result.get("doi") and not doi:
                df.at[idx, "doi"] = result["doi"]

        # 策略3: 用标题查 Semantic Scholar
        if title:
            result = _query_semantic_scholar(title, ss_cfg)
            if result and result.get("email"):
                _fill_result(df, idx, result, "SemanticScholar")
                continue

        df.at[idx, "lookup_status"] = "未找到邮箱"
        time.sleep(1)  # 礼貌性延迟

    return df


def _fill_result(df: pd.DataFrame, idx: int, result: dict, source: str):
    """将检索结果填入 DataFrame"""
    df.at[idx, "email"] = result.get("email", "")
    df.at[idx, "corresponding_author"] = result.get("corresponding_author", "")
    df.at[idx, "lookup_source"] = source
    df.at[idx, "lookup_status"] = "成功"
    if result.get("doi"):
        df.at[idx, "doi"] = result["doi"]
    print(f"    ✓ 找到邮箱: {result['email']} (via {source})")


def _query_crossref_by_doi(doi: str, config: dict) -> dict | None:
    """通过 DOI 查询 CrossRef"""
    base_url = config.get("base_url", "https://api.crossref.org")
    mailto = config.get("mailto", "")
    timeout = config.get("timeout", 15)

    url = f"{base_url}/works/{urllib.parse.quote(doi, safe='')}"
    params = {}
    if mailto:
        params["mailto"] = mailto

    try:
        resp = requests.get(url, params=params, timeout=timeout)
        if resp.status_code != 200:
            return None
        data = resp.json().get("message", {})
        return _extract_crossref_author_info(data)
    except Exception as e:
        print(f"    ✗ CrossRef DOI查询失败: {e}")
        return None


def _query_crossref_by_title(title: str, config: dict) -> dict | None:
    """通过标题查询 CrossRef"""
    base_url = config.get("base_url", "https://api.crossref.org")
    mailto = config.get("mailto", "")
    timeout = config.get("timeout", 15)

    url = f"{base_url}/works"
    params = {
        "query.title": title,
        "rows": 3,
        "select": "DOI,title,author",
    }
    if mailto:
        params["mailto"] = mailto

    try:
        resp = requests.get(url, params=params, timeout=timeout)
        if resp.status_code != 200:
            return None
        items = resp.json().get("message", {}).get("items", [])
        if not items:
            return None

        # 取第一个结果（CrossRef 已按相关性排序）
        best = items[0]
        # 简单验证标题是否匹配
        cr_title = " ".join(best.get("title", []))
        if not _titles_similar(title, cr_title):
            return None

        result = _extract_crossref_author_info(best)
        result["doi"] = best.get("DOI", "")
        return result
    except Exception as e:
        print(f"    ✗ CrossRef 标题查询失败: {e}")
        return None


def _extract_crossref_author_info(data: dict) -> dict:
    """从 CrossRef 返回数据中提取通讯作者信息"""
    result = {
        "email": "",
        "corresponding_author": "",
        "doi": data.get("DOI", ""),
    }

    authors = data.get("author", [])
    if not authors:
        return result

    # 优先查找标记了 ORCID 或有 affiliation 的通讯作者
    for author in authors:
        # CrossRef 中部分记录有 email 字段（少见但存在）
        if "email" in author:
            result["email"] = author["email"]
            name = f"{author.get('given', '')} {author.get('family', '')}".strip()
            result["corresponding_author"] = name
            return result

    # 如果没有直接的 email，记录第一作者姓名
    first = authors[0]
    result["corresponding_author"] = f"{first.get('given', '')} {first.get('family', '')}".strip()
    return result


def _query_semantic_scholar(title: str, config: dict) -> dict | None:
    """通过 Semantic Scholar API 查询"""
    base_url = config.get("base_url", "https://api.semanticscholar.org/graph/v1")
    api_key = config.get("api_key", "")
    timeout = config.get("timeout", 15)

    url = f"{base_url}/paper/search"
    params = {
        "query": title,
        "limit": 3,
        "fields": "title,authors,externalIds",
    }
    headers = {}
    if api_key:
        headers["x-api-key"] = api_key

    try:
        resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        if resp.status_code != 200:
            return None
        papers = resp.json().get("data", [])
        if not papers:
            return None

        best = papers[0]
        ss_title = best.get("title", "")
        if not _titles_similar(title, ss_title):
            return None

        # Semantic Scholar 通常不直接提供邮箱，但可以获取作者信息
        authors = best.get("authors", [])
        doi = best.get("externalIds", {}).get("DOI", "")

        result = {
            "email": "",
            "corresponding_author": authors[0]["name"] if authors else "",
            "doi": doi,
        }

        # 尝试通过作者详情页获取邮箱（部分作者公开了邮箱）
        if authors:
            author_id = authors[0].get("authorId")
            if author_id:
                email = _get_ss_author_email(author_id, base_url, headers, timeout)
                if email:
                    result["email"] = email

        return result
    except Exception as e:
        print(f"    ✗ Semantic Scholar 查询失败: {e}")
        return None


def _get_ss_author_email(author_id: str, base_url: str, headers: dict, timeout: int) -> str:
    """尝试从 Semantic Scholar 作者详情获取邮箱"""
    try:
        url = f"{base_url}/author/{author_id}"
        params = {"fields": "name,url,homepage"}
        resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        if resp.status_code == 200:
            # Semantic Scholar 不直接暴露邮箱，但 homepage 可能有用
            pass
        return ""
    except Exception:
        return ""


def _titles_similar(title1: str, title2: str) -> bool:
    """简单判断两个标题是否相似"""
    # 清理并比较
    def clean(t):
        t = t.lower().strip()
        t = "".join(c for c in t if c.isalnum() or c == " " or '\u4e00' <= c <= '\u9fff')
        return t

    c1, c2 = clean(title1), clean(title2)
    if not c1 or not c2:
        return False

    # 短标题用包含关系，长标题用字符重叠率
    shorter = min(c1, c2, key=len)
    longer = max(c1, c2, key=len)

    if len(shorter) <= 10:
        return shorter in longer

    # 计算字符级别的相似度
    set1, set2 = set(c1.split()), set(c2.split())
    if not set1 or not set2:
        return c1 in c2 or c2 in c1
    overlap = len(set1 & set2) / max(len(set1), len(set2))
    return overlap > 0.5


def save_results(df: pd.DataFrame, emails_csv: str, failed_csv: str):
    """保存检索结果：成功的和失败的分别存储"""
    os.makedirs(os.path.dirname(emails_csv), exist_ok=True)

    found = df[df["email"] != ""].copy()
    not_found = df[df["email"] == ""].copy()

    found.to_csv(emails_csv, index=False, encoding="utf-8-sig")
    not_found.to_csv(failed_csv, index=False, encoding="utf-8-sig")

    print(f"\n检索完成:")
    print(f"  ✓ 找到邮箱: {len(found)} 条 → {emails_csv}")
    print(f"  ✗ 未找到邮箱: {len(not_found)} 条 → {failed_csv}")
    print(f"  （请人工在知网查找未找到邮箱的条目，填入 manual_supplement.csv）")

    return found, not_found


def merge_manual_supplement(emails_csv: str, supplement_csv: str) -> pd.DataFrame:
    """将人工补充的邮箱数据合并到主数据"""
    df_main = pd.read_csv(emails_csv, encoding="utf-8-sig")

    if not os.path.exists(supplement_csv):
        print(f"未找到人工补充文件: {supplement_csv}")
        return df_main

    df_supp = pd.read_csv(supplement_csv, encoding="utf-8-sig")
    print(f"加载 {len(df_supp)} 条人工补充记录")

    # 补充文件至少需要 title 和 email 两列
    required_cols = {"title", "email"}
    if not required_cols.issubset(set(df_supp.columns)):
        print(f"错误：补充文件需要包含以下列: {required_cols}")
        return df_main

    # 合并：将补充数据追加到主数据
    new_rows = []
    for _, row in df_supp.iterrows():
        if pd.notna(row["email"]) and str(row["email"]).strip():
            new_rows.append(row.to_dict())

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_merged = pd.concat([df_main, df_new], ignore_index=True)
        df_merged.to_csv(emails_csv, index=False, encoding="utf-8-sig")
        print(f"✓ 合并完成，总计 {len(df_merged)} 条有邮箱的记录")
        return df_merged

    print("补充文件中没有有效的邮箱记录")
    return df_main
