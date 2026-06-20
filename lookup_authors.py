"""
第二阶段：通过本地 Excel 数据库检索被引作者的联系邮箱
"""

import os

import pandas as pd

from db_lookup import AuthorDatabase, create_database


def lookup_emails(references_csv: str, config: dict, author_db: AuthorDatabase = None) -> pd.DataFrame:
    """
    读取参考文献 CSV，逐条通过本地数据库检索作者邮箱。
    同一篇参考文献有多个作者时，每个找到邮箱的作者单独生成一行。
    返回包含邮箱信息的 DataFrame。
    """
    if author_db is None:
        db_dir = config.get("paths", {}).get("local_db_dir", "数据")
        if os.path.exists(db_dir):
            author_db = create_database(db_dir)
        else:
            print(f"[警告] 本地数据库目录不存在: {db_dir}，所有记录将标记为未找到邮箱")

    df = pd.read_csv(references_csv, encoding="utf-8-sig")
    print(f"共加载 {len(df)} 条参考文献记录")

    result_rows = []

    for idx, row in df.iterrows():
        authors = str(row.get("authors", "")).strip()
        title = str(row.get("title", "")).strip()

        if not authors:
            new_row = row.to_dict()
            new_row.update({"corresponding_author": "", "email": "", "matched_name": "",
                            "lookup_source": "", "lookup_status": "跳过：作者字段为空"})
            result_rows.append(new_row)
            continue

        print(f"  [{idx + 1}/{len(df)}] 检索: {title[:50]}...")

        found_any = False
        if author_db:
            for author_name in _parse_author_names(authors):
                db_result = author_db.lookup_by_name(author_name)
                if db_result and db_result.get("email"):
                    new_row = row.to_dict()
                    new_row["email"] = db_result.get("email", "")
                    new_row["corresponding_author"] = db_result.get("corresponding_author", "") or db_result.get("name", "")
                    new_row["matched_name"] = author_name
                    new_row["lookup_source"] = "本地数据库"
                    new_row["lookup_status"] = "成功"
                    result_rows.append(new_row)
                    print(f"    ✓ 找到邮箱: {db_result['email']} (匹配姓名: {author_name})")
                    found_any = True

        if not found_any:
            new_row = row.to_dict()
            new_row.update({"corresponding_author": "", "email": "", "matched_name": "",
                            "lookup_source": "", "lookup_status": "未找到邮箱"})
            result_rows.append(new_row)

    result_df = pd.DataFrame(result_rows)
    for col in ["corresponding_author", "email", "matched_name", "lookup_source", "lookup_status"]:
        if col not in result_df.columns:
            result_df[col] = ""
    return result_df


def _parse_author_names(authors_str: str) -> list:
    """将 authors 字段拆分为姓名列表，支持分号或逗号分隔"""
    if not authors_str or authors_str.strip() in ("", "nan"):
        return []
    if ";" in authors_str:
        names = [n.strip() for n in authors_str.split(";")]
    else:
        names = [n.strip() for n in authors_str.split(",")]
    # 过滤无意义的占位词
    skip_tokens = {"等", "et al", "et al.", "others"}
    seen = set()
    result = []
    for name in names:
        if name and name.lower() not in skip_tokens and name not in seen:
            seen.add(name)
            result.append(name)
    return result


def save_results(df: pd.DataFrame, emails_csv: str, failed_csv: str):
    """保存检索结果：成功的和失败的分别存储"""
    dir_name = os.path.dirname(emails_csv)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)

    found = df[df["email"] != ""].copy()
    not_found = df[df["email"] == ""].copy()

    found.to_csv(emails_csv, index=False, encoding="utf-8-sig")
    not_found.to_csv(failed_csv, index=False, encoding="utf-8-sig")

    print(f"\n检索完成:")
    print(f"  ✓ 找到邮箱: {len(found)} 条 → {emails_csv}")
    print(f"  ✗ 未找到邮箱: {len(not_found)} 条 → {failed_csv}")

    return found, not_found


def merge_manual_supplement(emails_csv: str, supplement_csv: str) -> pd.DataFrame:
    """将人工补充的邮箱数据合并到主数据"""
    df_main = pd.read_csv(emails_csv, encoding="utf-8-sig")

    if not os.path.exists(supplement_csv):
        print(f"未找到人工补充文件: {supplement_csv}")
        return df_main

    df_supp = pd.read_csv(supplement_csv, encoding="utf-8-sig")
    print(f"加载 {len(df_supp)} 条人工补充记录")

    required_cols = {"title", "email"}
    if not required_cols.issubset(set(df_supp.columns)):
        print(f"错误：补充文件需要包含以下列: {required_cols}")
        return df_main

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
