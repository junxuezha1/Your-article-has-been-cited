"""
db_lookup.py - Local Excel database lookup module for author email queries.
Loads local `.xls` author databases and queries them by name.
"""

import os
import warnings
from difflib import SequenceMatcher
from typing import Optional

import pandas as pd


class AuthorDatabase:
    """Loads and queries local Excel author databases."""

    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        self.df: pd.DataFrame = pd.DataFrame()
        self.load()

    def load(self):
        """Load all xls files from data_dir into a unified DataFrame."""
        frames = []

        # --- Format A: persons (1).xls, persons (2).xls, persons (3).xls ---
        # Row 0 = title row "人员列表" (skip), Row 1 = header: 姓名, 性别, 通讯地址, 电话, 邮政编码, 邮箱, 学科专业, 手机号码, ...
        persons_files = [
            "persons (1).xls",
            "persons (2).xls",
            "persons (3).xls",
        ]
        for fname in persons_files:
            fpath = os.path.join(self.data_dir, fname)
            if not os.path.exists(fpath):
                print(f"[db_lookup] Warning: file not found, skipping: {fpath}")
                continue
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    raw = pd.read_excel(fpath, header=1, engine="xlrd", dtype=str)
                # Normalise column names (strip whitespace)
                raw.columns = [str(c).strip() for c in raw.columns]
                frame = pd.DataFrame({
                    "name":        raw.get("姓名"),
                    "email":       raw.get("邮箱"),
                    "address":     raw.get("通讯地址"),
                    "phone":       raw.get("手机号码").fillna(raw.get("电话", ""))
                                   if "手机号码" in raw.columns else raw.get("电话"),
                    "institution": raw.get("学科专业"),
                })
                frames.append(frame)
                print(f"[db_lookup] Loaded {len(frame)} rows from {fname}")
            except Exception as exc:
                print(f"[db_lookup] Warning: could not read {fpath}: {exc}")

        # --- Format B: 20260401115600.xls ---
        # Row 0 = description text (skip), Row 1 = header, data from Row 2
        format_b_file = "20260401115600.xls"
        fpath_b = os.path.join(self.data_dir, format_b_file)
        if not os.path.exists(fpath_b):
            print(f"[db_lookup] Warning: file not found, skipping: {fpath_b}")
        else:
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    raw = pd.read_excel(
                        fpath_b, header=1, skiprows=[0], engine="xlrd", dtype=str
                    )
                raw.columns = [str(c).strip() for c in raw.columns]
                frame = pd.DataFrame({
                    "name":        raw.get("姓名（中文）"),
                    "email":       raw.get("E-mail"),
                    "address":     raw.get("通信地址"),
                    "phone":       raw.get("手机号"),
                    "institution": raw.get("单位名称"),
                })
                frames.append(frame)
                print(f"[db_lookup] Loaded {len(frame)} rows from {format_b_file}")
            except Exception as exc:
                print(f"[db_lookup] Warning: could not read {fpath_b}: {exc}")

        if frames:
            self.df = pd.concat(frames, ignore_index=True)
            # Drop rows where both name and email are missing
            self.df.dropna(subset=["name", "email"], how="all", inplace=True)
            # Strip whitespace from string columns
            for col in self.df.columns:
                self.df[col] = self.df[col].astype(str).str.strip().replace("nan", "")
            print(f"[db_lookup] Total records loaded: {len(self.df)}")
        else:
            print("[db_lookup] Warning: no data files were loaded.")

    def _row_to_dict(self, row: pd.Series) -> dict:
        return {
            "name":        row.get("name", ""),
            "email":       row.get("email", ""),
            "address":     row.get("address", ""),
            "phone":       row.get("phone", ""),
            "institution": row.get("institution", ""),
        }

    def lookup_by_name(self, name: str) -> Optional[dict]:
        """
        Look up an author by name.
        First tries exact match, then substring containment.
        Returns the first match as a dict, or None.
        """
        if self.df.empty or not name:
            return None

        name_stripped = name.strip()

        # 1. Exact match
        mask_exact = self.df["name"] == name_stripped
        if mask_exact.any():
            return self._row_to_dict(self.df[mask_exact].iloc[0])

        # 2. Substring containment (name in cell OR cell in name)
        mask_sub = self.df["name"].str.contains(name_stripped, na=False, regex=False)
        if mask_sub.any():
            return self._row_to_dict(self.df[mask_sub].iloc[0])

        return None

    def lookup_by_name_fuzzy(
        self, name: str, threshold: float = 0.8
    ) -> list[dict]:
        """
        Return all records whose name similarity to `name` exceeds `threshold`.
        Uses SequenceMatcher ratio for fuzzy matching.
        """
        if self.df.empty or not name:
            return []

        name_stripped = name.strip()
        results = []

        for _, row in self.df.iterrows():
            cell_name = row.get("name", "")
            if not cell_name:
                continue
            ratio = SequenceMatcher(None, name_stripped, cell_name).ratio()
            if ratio >= threshold:
                entry = self._row_to_dict(row)
                entry["_similarity"] = round(ratio, 4)
                results.append(entry)

        # Sort by similarity descending
        results.sort(key=lambda x: x["_similarity"], reverse=True)
        return results


def create_database(data_dir: str) -> "AuthorDatabase":
    """Create and return a loaded AuthorDatabase instance."""
    return AuthorDatabase(data_dir)


if __name__ == "__main__":
    import sys

    # Default data dir: '数据' subfolder next to this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_data_dir = os.path.join(script_dir, "数据")
    data_dir = sys.argv[1] if len(sys.argv) > 1 else default_data_dir

    print(f"Loading database from: {data_dir}")
    db = create_database(data_dir)
    print(f"\nTotal records in database: {len(db.df)}")

    # Quick smoke test
    if len(db.df) > 0:
        sample_name = db.df["name"].dropna().iloc[0]
        print(f"\nSample lookup for name: '{sample_name}'")
        result = db.lookup_by_name(sample_name)
        print(f"  Result: {result}")
