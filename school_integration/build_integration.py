"""学校規模 × 競合情報を管理表の行ごとに統合し、CSV / JSON に書き出す。

    python3 build_integration.py 管理表.xlsx tabs.json match.json size_overrides.json OUT_DIR [--cutoff 2026-09-30]

出力
  OUT_DIR/integration.json  … 行ごとの統合結果（根拠つき）
  OUT_DIR/integration.csv   … スプレッドシート取り込み用（管理表の行番号つき）
"""
import argparse
import csv
import json
import os
import re
import unicodedata

import openpyxl

from competitors import AGIC, NOT_HEARD, classify_text, find_mentions
from match_rows import SHEET, norm as school_norm, score as name_score
from school_size import normalize as size_normalize, prep as size_prep

OUT_COLUMNS = [
    "行", "学校名", "訪問者", "学校規模", "AGIC訪問", "他社訪問", "訪問競合社名",
    "情報時点", "根拠（学校規模）", "根拠（競合）", "抽出メモ",
]


def nfkc(s):
    return unicodedata.normalize("NFKC", str(s or ""))


def short(s, n=90):
    s = re.sub(r"\s+", " ", nfkc(s)).strip()
    return s if len(s) <= n else s[: n - 1] + "…"


def same_school(a, b):
    """既存の競合情報との突合は厳しめに判定する（「日本大学」が全学部に波及しないように）"""
    na, nb = school_norm(a), school_norm(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    if na in nb or nb in na:
        return min(len(na), len(nb)) / max(len(na), len(nb)) >= 0.85
    import difflib
    return difflib.SequenceMatcher(None, na, nb).ratio() >= 0.88


# 競合の「来ている／来ていない」を答える欄（他の欄の「特になし」は競合の有無ではない）
PRESENCE_LABELS = ("BPOや新卒派遣の引き合い", "派遣会社はきてるか", "よく来ている競合他社（その他欄）",
                   "上記他社の27卒の案内はいつから受けたか")


# ---------------------------------------------------------------- 既存の競合情報（管理表内）
PRESENT_WORDS = re.compile(r"^(接点あり|あり|複数あり|多数あり|数社あり|数十社あり|\d+[～〜~-]?\d*社(ほど)?あり|過去1社あり|専門職はあり)")
NONE_WORDS = re.compile(r"^(接点なし|なし|訪問はなし|ほぼない|ほぼなし|ほぼ来ていない|あまりない)")


def prior_status(v):
    v = nfkc(v).strip()
    if not v:
        return None
    if PRESENT_WORDS.match(v):
        return "present"
    if NONE_WORDS.match(v):
        return "none"
    return "unknown"


def load_prior_sources(wb):
    """前回までの競合情報。{school, src, date, status(present/none/unknown/None), names, text}"""
    prior = []
    ws = wb["他社集約用"]  # 2〜4月訪問分の集約（R:接点の有無, S:競合他社名）
    for row in ws.iter_rows(min_row=2, values_only=True):
        school, status, names = row[0], row[17], row[18]
        if not school or not (status or names):
            continue
        prior.append({"school": school, "src": "他社集約用", "date": "",
                      "status": prior_status(status) or ("present" if names else None),
                      "names": nfkc(names or ""), "text": f"{status or ''} {names or ''}".strip()})
    ws = wb["訪問学校の競合調査"]  # 前回の競合調査（E:競合の訪問状況, F:判明している競合社名）
    for row in ws.iter_rows(min_row=3, values_only=True):
        school = row[2]
        if not school:
            continue
        d = row[3]
        date = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else nfkc(d or "")
        names = "" if nfkc(row[5] or "").strip() in ("―", "-", "") else nfkc(row[5])
        prior.append({"school": school, "src": "前回競合調査", "date": date,
                      "status": prior_status(row[4]), "names": names,
                      "text": " ".join(str(v) for v in (row[4], row[5], row[7]) if v not in (None, "―", "-"))})
    ws = wb[SHEET]  # 27卒タブ AJ列：行がずれているため【学校名】見出しで学校を特定し直す
    for r in range(2, ws.max_row + 1):
        v = ws[f"AJ{r}"].value
        if not v:
            continue
        for part in re.split(r"(?=【)", str(v)):
            m = re.match(r"【([^】]+)】\s*([\d/・、,]*)\s*(?:訪問)?(.*)", part, flags=re.S)
            if m:
                prior.append({"school": m.group(1).strip(), "src": "AJ列", "date": "",
                              "status": None, "names": "", "text": m.group(3)})
    return prior


# ---------------------------------------------------------------- 学校規模
def decide_size(row_no, blocks, overrides):
    key = f"row:{row_no}"
    fix = overrides.get(f"rowfix:{row_no}")
    if fix:  # 複数回の訪問の記述を合わせて読む必要がある行（最新値だけでは誤解を招くもの）
        return fix, "memo", "訪問シートの複数回の記述から整理（行ごとの補正）"
    dated = sorted(blocks, key=lambda b: b["date"] or "0000", reverse=True)
    for b in dated:
        v, kind = size_normalize(b["size_raw"], overrides)
        if v:
            return v, kind, f"訪問シート({b['date'] or '日付不明'}) 最上級生人数欄：{short(b['size_raw'], 60)}"
    if key in overrides and overrides[key]:
        return overrides[key], "memo", "訪問シートのメモ本文から推定（人数欄は空欄）"
    raws = [b["size_raw"] for b in dated if b["size_raw"]]
    if raws:
        v, kind = size_normalize(raws[0], overrides)
        return None, "review", f"人数欄の記述を数値化できず：{short(raws[0], 60)}"
    return None, "empty", "訪問シートに人数の記載なし"


# ---------------------------------------------------------------- 競合
def decide_competitors(blocks, prior_hits, cutoff):
    evid = []        # (日付, 出典, 抜粋)
    signals = set()  # "present" / "none"（情報の食い違い検出用）
    agic = None      # True / False / None
    others = {}      # 社名 -> True/False
    other_unknown_present = False
    none_all = False
    heard_nothing = True
    latest = ""

    def absorb(text, date, source, presence=True):
        nonlocal agic, other_unknown_present, none_all, heard_nothing, latest
        comps, none_flag, unk, not_heard = classify_text(text)
        if not presence:
            none_flag = False
        meaningful = bool(comps) or none_flag or unk
        if meaningful:
            heard_nothing = False
            latest = max(latest, date or "")
        for c, pos in comps.items():
            if c == AGIC:
                agic = True if (pos or agic) else (agic if agic is not None else False)
            else:
                others[c] = others.get(c, False) or pos
        if none_flag and not comps:
            none_all = True
            signals.add("none")
        if unk or any(comps.values()):
            signals.add("present")
        if unk:
            other_unknown_present = True
        if meaningful:
            snips = [s for c, _, s in find_mentions(text)]
            evid.append((date or "", source, short(snips[0] if snips else text, 70)))
        elif not_heard and text.strip():
            evid.append((date or "", source, short(text, 50)))

    for b in blocks:
        if b["date"] and b["date"] > cutoff:
            continue
        pres = "\n".join(v for k, v in b["competitor"].items() if k in PRESENCE_LABELS)
        rest = "\n".join(v for k, v in b["competitor"].items() if k not in PRESENCE_LABELS)
        if pres:
            absorb(pres, b["date"], "訪問シート競合欄")
        if rest and find_mentions(rest):
            absorb(rest, b["date"], "訪問シート競合欄", presence=False)
        # メモ本文は社名の言及があるときだけ拾う（雑音を避ける）
        if find_mentions(b["memo"]):
            absorb(b["memo"], b["date"], "訪問シートメモ")
    for p in prior_hits:
        if p["status"] == "present":
            heard_nothing = False
            signals.add("present")
            comps, _, _, _ = classify_text(p["names"])
            for c in comps:
                if c == AGIC:
                    agic = True
                else:
                    others[c] = True
            if not [c for c in comps if c != AGIC] and not comps:
                other_unknown_present = True
            evid.append((p["date"], p["src"], short(p["text"], 70)))
        elif p["status"] == "none":
            heard_nothing = False
            none_all = True
            signals.add("none")
            evid.append((p["date"], p["src"], short(p["text"], 70)))
        elif p["status"] == "unknown":
            evid.append((p["date"], p["src"], short(p["text"], 50)))
        else:
            absorb(p["text"], p["date"], p["src"])

    pos_others = sorted(c for c, p in others.items() if p)
    any_presence = bool(pos_others) or other_unknown_present
    if agic is True:
        agic_v = "あり"
    elif agic is False:
        agic_v = "なし"          # エイジェックについて明示的に「来ていない」
    elif none_all and not any_presence:
        agic_v = "なし"          # 人材系・競合の訪問自体が無い
    else:
        agic_v = "不明"
    if pos_others:
        other_v = "あり"
    elif other_unknown_present:
        other_v = "あり"
    elif none_all or (others and not pos_others):
        other_v = "なし"
    else:
        other_v = "不明"
    names = ", ".join(pos_others)
    if other_v == "あり" and not names:
        names = "社名不明"
    evid.sort(key=lambda e: e[0], reverse=True)
    ev = " / ".join(f"{s}({d or '日付不明'})：{x}" for d, s, x in evid[:2])
    if heard_nothing and not evid:
        ev = "競合についての記載なし"
    conflict = signals == {"present", "none"}
    return agic_v, other_v, names, latest, ev, conflict


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("tabs")
    ap.add_argument("match")
    ap.add_argument("size_overrides")
    ap.add_argument("out_dir")
    ap.add_argument("--cutoff", default="2026-09-30")
    a = ap.parse_args()

    tabs = json.load(open(a.tabs, encoding="utf-8"))
    match = json.load(open(a.match, encoding="utf-8"))
    overrides = json.load(open(a.size_overrides, encoding="utf-8"))
    wb = openpyxl.load_workbook(a.xlsx)
    prior = load_prior_sources(wb)

    records = []
    for m in match:
        blocks = [b for i in m["tab_idx"] for b in tabs[i]["blocks"]]
        prior_hits = [p for p in prior if same_school(m["school"], p["school"])]
        size, size_kind, size_ev = decide_size(m["row"], blocks, overrides)
        agic, other, names, latest, comp_ev, conflict = decide_competitors(blocks, prior_hits, a.cutoff)
        notes = []
        if not m["tab_idx"]:
            notes.append("訪問シート未接続" if not m["doc_accessible"] else "訪問シートなし")
        elif m.get("score", 1) < 0.95 and not m.get("manual"):
            notes.append(f"シート照合:{m['best_tab']}")
        if size_kind in ("review", "memo"):
            notes.append("規模要確認")
        if conflict:
            notes.append("競合情報が食い違い（新旧で有無が異なる）")
        records.append({
            "行": m["row"], "学校名": m["school"], "訪問者": m["visitor"] or "",
            "学校規模": size or "", "AGIC訪問": agic, "他社訪問": other, "訪問競合社名": names,
            "情報時点": latest, "根拠（学校規模）": size_ev, "根拠（競合）": comp_ev,
            "抽出メモ": " / ".join(notes),
        })

    os.makedirs(a.out_dir, exist_ok=True)
    with open(os.path.join(a.out_dir, "integration.json"), "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=1)
    with open(os.path.join(a.out_dir, "integration.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=OUT_COLUMNS)
        w.writeheader()
        w.writerows(records)
    print(f"{len(records)} rows -> {a.out_dir}")


if __name__ == "__main__":
    main()
