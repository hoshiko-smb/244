"""管理表「27卒_アポイント＆訪問管理」の各行を、訪問シートのタブ（学校）に突き合わせる。

    python3 match_rows.py 管理表.xlsx tabs.json OUT_match.json

突き合わせの優先順位
 1. AB列（訪問シート）の URL が指すドキュメント内のタブで、学校名が一致/包含するもの
 2. 全ドキュメントのタブから学校名が一致/包含するもの
 3. 文字列類似度（difflib）が高いもの（要確認フラグ付き）
"""
import difflib
import json
import re
import sys
import unicodedata

import openpyxl

SHEET = "27卒_アポイント＆訪問管理"

DROP_WORDS = [
    "総合学園", "総合学院", "学校法人", "専門学校", "専門職大学", "キャンパス", "校舎",
]


SCHOOL_SUFFIX = r"(大学|短期大学|短期大学部|学院|学園|校|カレッジ|アカデミー|スクール)$"
OHARA_CAMPUS = ["津田沼", "立川", "町田", "大宮", "千葉", "柏", "横浜", "池袋", "東京", "水道橋", "難波"]


def ohara_key(s: str) -> str | None:
    """大原系の正式名称（例: 大原簿記公務員専門学校 千葉校）を訪問シートのタブ名（大原学園 千葉校）に寄せる"""
    s = unicodedata.normalize("NFKC", str(s or ""))
    if not s.startswith("大原"):
        return None
    campus = next((c for c in OHARA_CAMPUS if c + "校" in s or c + " " in s or s.endswith(c)), None)
    if not campus:
        return None
    sub = ""
    if "IT" in s:
        sub = "IT"
    elif "簿記" in s and "柏" in s:
        sub = "簿記"
    elif "ビジネス" in s and campus == "横浜":
        sub = "ビジネス"
    return f"大原学園{campus}校{sub}"


def norm(s: str) -> str:
    k = ohara_key(s)
    if k:
        return k
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = s.replace("アンド", "&")
    s = re.sub(r"[（(【\[].*?[）)】\]]", "", s)  # 括弧内の補足は除去
    s = s.replace("＆", "&").replace("・", "").replace("･", "")
    s = re.sub(r"\s+", "", s)
    for w in DROP_WORDS:
        s = s.replace(w, "")
    return s


def doc_id_of(cell) -> str | None:
    txt = str(cell.value or "")
    if cell.hyperlink is not None and cell.hyperlink.target:
        txt += " " + cell.hyperlink.target
    m = re.search(r"document/d/([\w-]+)", txt)
    return m.group(1) if m else None


def tab_score(n: str, t: dict) -> float:
    # タブ名を優先し、シート内「学校名」欄は補助（少し割り引く）
    alt = max((score(n, f) for f in t["norm_alt"]), default=0.0) * 0.97
    return max(score(n, t["norm"]), alt)


def score(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        sc = 0.9 * min(len(a), len(b)) / max(len(a), len(b)) + 0.1
        # 「東京福祉大学」と「東京福祉大学池袋」のように、学校名＋キャンパス名の違いだけなら同一校扱い
        short = a if len(a) <= len(b) else b
        if (a.startswith(b) or b.startswith(a)) and len(short) >= 4 and re.search(SCHOOL_SUFFIX, short):
            sc = max(sc, 0.8)
        return sc
    return difflib.SequenceMatcher(None, a, b).ratio() * 0.85


def main():
    xlsx, tabs_json, out = sys.argv[1:4]
    manual = json.load(open(sys.argv[4], encoding="utf-8")) if len(sys.argv) > 4 else {}
    tabs = json.load(open(tabs_json, encoding="utf-8"))
    for i, t in enumerate(tabs):
        t["idx"] = i
        t["norm"] = norm(t["tab"])
        t["norm_alt"] = [norm(f) for f in t.get("school_fields", [])]
    wb = openpyxl.load_workbook(xlsx, read_only=False)
    ws = wb[SHEET]
    res = []
    for r in range(2, ws.max_row + 1):
        school = ws[f"H{r}"].value
        if not school:
            continue
        doc = doc_id_of(ws[f"AB{r}"])
        n = norm(school)
        key = str(r)
        if key in manual or school in manual:
            target = manual.get(key, manual.get(school))
            idx = [t["idx"] for t in tabs if target and t["tab"] == target]
            res.append({"row": r, "school": school, "visitor": ws[f"T{r}"].value or ws[f"D{r}"].value,
                        "doc_id": doc, "doc_accessible": (doc is None) or any(t["doc_id"] == doc for t in tabs),
                        "score": 1.0 if idx else 0.0, "tab_idx": idx, "best_tab": target, "manual": True})
            continue
        pool = [t for t in tabs if doc and t["doc_id"] == doc] or tabs
        ranked = sorted(((tab_score(n, t), t) for t in pool), key=lambda x: (-x[0], x[1]["tab"], x[1]["doc_id"]))
        if doc and (not ranked or ranked[0][0] < 0.75):
            # URL 先のドキュメントに見当たらない場合は全ドキュメントから探す
            ranked = sorted(((tab_score(n, t), t) for t in tabs), key=lambda x: (-x[0], x[1]["tab"], x[1]["doc_id"]))
        best = ranked[0] if ranked else (0, None)
        # 同名タブ（同じ学校の別キャンパス/重複タブ）は全て候補に残す
        same = [t["idx"] for s, t in ranked if s >= max(0.75, best[0] - 1e-9)]
        # 訪問履歴が別のシート（年度ごとの No.1〜No.4）に続いている場合も拾う：学校名がほぼ完全一致するタブは合算
        if best[0] >= 0.75:
            for t in tabs:
                if t["idx"] not in same and score(n, t["norm"]) >= 0.95:
                    same.append(t["idx"])
        res.append({
            "row": r,
            "school": school,
            "visitor": ws[f"T{r}"].value or ws[f"D{r}"].value,
            "doc_id": doc,
            "doc_accessible": (doc is None) or any(t["doc_id"] == doc for t in tabs),
            "score": round(best[0], 3),
            "tab_idx": same if best[0] >= 0.75 else [],
            "best_tab": best[1]["tab"] if best[1] else None,
        })
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    ok = sum(1 for x in res if x["tab_idx"])
    print(f"rows={len(res)} matched={ok} unmatched={len(res) - ok}")


if __name__ == "__main__":
    main()
