"""統合結果から、エリア × AGIC訪問 × 学校規模 の競争状況サマリーと、攻略候補リストを作る。

    python3 summarize.py 管理表.xlsx integration.json OUT_DIR
出力: OUT_DIR/competition_summary.md, OUT_DIR/agic_targets.csv
"""
import csv
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict

import openpyxl

SHEET = "27卒_アポイント＆訪問管理"
WARDS = ["千代田区", "中央区", "港区", "新宿区", "文京区", "台東区", "墨田区", "江東区", "品川区", "目黒区",
         "大田区", "世田谷区", "渋谷区", "中野区", "杉並区", "豊島区", "北区", "荒川区", "板橋区", "練馬区",
         "足立区", "葛飾区", "江戸川区"]
CENTER = {"千代田区", "中央区", "港区", "新宿区", "渋谷区"}  # 都心5区
AREAS = ["都心5区", "23区（都心5区以外）", "東京都（23区外）", "神奈川", "埼玉", "千葉", "関西", "その他地域",
         "WEB面談・住所未記入"]
BANDS = ["300人以上", "100〜299人", "100人未満", "規模不明"]


TOKYO_CITIES = ["八王子市", "町田市", "立川市", "武蔵野市", "三鷹市", "小金井市", "国分寺市", "多摩市", "調布市",
                "府中市", "日野市", "小平市", "東村山市", "西東京市", "国立市", "狛江市", "稲城市", "青梅市", "昭島市"]
PREF_CITIES = {
    "神奈川": ["横浜市", "川崎市", "相模原市", "藤沢市", "鎌倉市", "厚木市", "平塚市", "小田原市", "綾瀬市", "横須賀市"],
    "埼玉": ["さいたま市", "川口市", "川越市", "所沢市", "越谷市", "坂戸市", "上尾市", "春日部市", "草加市"],
    "千葉": ["千葉市", "船橋市", "柏市", "松戸市", "習志野市", "浦安市", "市川市", "成田市", "東金市"],
    "関西": ["大阪市", "京都市", "神戸市", "堺市", "姫路市", "西宮市", "芦屋市"],
}


def area_of(addr: str) -> str:
    a = unicodedata.normalize("NFKC", addr or "")
    a = re.sub(r"〒?\s*\d{3}-?\d{4}", "", a)
    a = re.sub(r"\s+", "", a)
    a = re.sub(r"^(WEB|Web|web|オンライン)[/／:：]?", "", a)
    if not a:
        return "WEB面談・住所未記入"
    ward = next((w for w in WARDS if a.startswith("東京都" + w) or a.startswith("東京" + w) or a.startswith(w)), None)
    if ward:
        return "都心5区" if ward in CENTER else "23区（都心5区以外）"
    if a.startswith("東京") or any(a.startswith(c) for c in TOKYO_CITIES):
        return "東京都（23区外）"
    for pref in ("神奈川", "埼玉", "千葉"):
        if a.startswith(pref):
            return pref
    if re.match(r"(大阪|京都|兵庫|奈良|滋賀|和歌山)", a):
        return "関西"
    for pref, cities in PREF_CITIES.items():
        if any(a.startswith(c) for c in cities):
            return pref
    return "その他地域"


def size_num(s: str):
    m = re.search(r"1学年([\d,]+)", s or "")
    return int(m.group(1).replace(",", "")) if m else None


def band(n):
    if n is None:
        return "規模不明"
    return "300人以上" if n >= 300 else "100〜299人" if n >= 100 else "100人未満"


def main():
    xlsx, integ, out = sys.argv[1:4]
    ws = openpyxl.load_workbook(xlsx, read_only=True)[SHEET]
    addr = {i: (row[10] or "") for i, row in enumerate(ws.iter_rows(values_only=True), start=1)}
    recs = json.load(open(integ, encoding="utf-8"))
    for r in recs:
        r["エリア"] = area_of(addr.get(r["行"], ""))
        r["_n"] = size_num(r["学校規模"])
        r["規模帯"] = band(r["_n"])

    lines = ["# 27卒 訪問校の競争状況（学校規模 × 競合）", "",
             f"対象: 27卒_アポイント＆訪問管理 の {len(recs)} 行／競合情報は 2026-09-30 までの訪問分", ""]

    lines += ["## エリア別 AGIC（エイジェック）訪問状況", "",
              "| エリア | 学校数 | AGICあり | AGICなし | AGIC不明 | 他社あり | 規模300人以上 |", "|---|--:|--:|--:|--:|--:|--:|"]
    for a in AREAS:
        rs = [r for r in recs if r["エリア"] == a]
        if not rs:
            continue
        c = Counter(r["AGIC訪問"] for r in rs)
        lines.append(f"| {a} | {len(rs)} | {c['あり']} | {c['なし']} | {c['不明']} | "
                     f"{sum(r['他社訪問'] == 'あり' for r in rs)} | {sum(r['規模帯'] == '300人以上' for r in rs)} |")
    lines.append("")

    tokyo = [r for r in recs if r["エリア"] in AREAS[:3]]
    lines += ["## 東京都内：規模帯 × AGIC訪問", "", "| 規模帯 | AGICあり | AGICなし | AGIC不明 | 計 |", "|---|--:|--:|--:|--:|"]
    for b in BANDS:
        rs = [r for r in tokyo if r["規模帯"] == b]
        c = Counter(r["AGIC訪問"] for r in rs)
        lines.append(f"| {b} | {c['あり']} | {c['なし']} | {c['不明']} | {len(rs)} |")
    lines.append("")

    comp = Counter(n for r in tokyo for n in r["訪問競合社名"].split(", ") if n and n != "社名不明")
    lines += ["## 東京都内で名前が挙がった競合（AGIC以外・学校数）", ""]
    lines += [f"- {k}: {v}校" for k, v in comp.most_common(12)]
    lines.append("")

    # 攻略候補：都内で AGIC が入っている学校（＝シェア奪取の主戦場）を規模の大きい順に
    def uniq(rows):  # 管理表に同じ学校が2行ある場合は1校として数える
        seen, out = set(), []
        for r in rows:
            if r["学校名"] not in seen:
                seen.add(r["学校名"])
                out.append(r)
        return out

    targets = uniq(sorted((r for r in tokyo if r["AGIC訪問"] == "あり"), key=lambda r: -(r["_n"] or 0)))
    hearing = uniq(sorted((r for r in tokyo if r["AGIC訪問"] == "不明" and (r["_n"] or 0) >= 300), key=lambda r: -(r["_n"] or 0)))
    lines += [f"## 攻略候補：都内で AGIC 訪問ありの学校（{len(targets)}校・規模順 上位20）", "",
              "| 学校名 | エリア | 学校規模 | 他社 | 訪問者 |", "|---|---|---|---|---|"]
    for r in targets[:20]:
        lines.append(f"| {r['学校名']} | {r['エリア']} | {r['学校規模'] or '不明'} | {r['訪問競合社名'] or '-'} | {r['訪問者']} |")
    lines += ["", f"## 要ヒアリング：都内・1学年300人以上で AGIC 不明の学校（{len(hearing)}校・上位15）", "",
              "| 学校名 | エリア | 学校規模 | 訪問者 |", "|---|---|---|---|"]
    for r in hearing[:15]:
        lines.append(f"| {r['学校名']} | {r['エリア']} | {r['学校規模']} | {r['訪問者']} |")
    lines.append("")

    os.makedirs(out, exist_ok=True)
    open(os.path.join(out, "competition_summary.md"), "w", encoding="utf-8").write("\n".join(lines))
    cols = ["行", "学校名", "エリア", "学校規模", "規模帯", "AGIC訪問", "他社訪問", "訪問競合社名", "訪問者", "根拠（競合）"]
    with open(os.path.join(out, "agic_targets.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in sorted(recs, key=lambda r: (AREAS.index(r["エリア"]), {"あり": 0, "不明": 1, "なし": 2}[r["AGIC訪問"]], -(r["_n"] or 0))):
            w.writerow(r)
    print("\n".join(lines[:40]))


if __name__ == "__main__":
    main()
