"""訪問シート（Google ドキュメント／タブ構成）から学校規模・競合情報を抽出する。

入力: Google Drive の read_file_content で取得したドキュメント本文（Markdown 風テキスト）。
      各タブは "# タブ名" 見出しで区切られている。
出力: タブ（=学校）ごとの訪問ブロック一覧を JSON で書き出す。

    python3 extract_visit_sheets.py OUT.json DOC_ID=path/to/doc.txt [DOC_ID=path ...]

AI に全文を読ませずに済むよう、テンプレートの決まった欄（最上級生の全体人数・競合情報欄）と
メモ欄の該当行だけを機械的に切り出す。判断が割れるものだけを後段で人（または AI）が確認する。
"""
import json
import re
import sys
import unicodedata

# 訪問シートの4系統のテンプレートで競合情報が入る欄
COMPETITOR_LABELS = [
    "BPOや新卒派遣の引き合い",            # ①標準テンプレート
    "AnewSとの違い",
    "どちらが魅力的か",
    "派遣会社はきてるか",                  # ③旧テンプレート（〜2024年度）
    "上記他社の27卒の案内はいつから受けたか",  # ②競合深掘りテンプレート
    "他社が説明会実施できて当社ができない理由",
]
# テンプレートの説明文（未記入のまま残っている定型文）は競合情報として扱わない
TEMPLATE_PLACEHOLDERS = [
    "※引き合い有の場合(企業名/業務内容/募集人数/強みなど)",
    "学生/親/学校目線など",
]
LABEL_WORDS = {"学生の傾向", "事務職への関心度・需要", "都内在住の割合/人数"}
# ②競合深掘りテンプレートの「★よく来ている競合他社」欄に印字されている7社の定型リスト。
# チェック状態はエクスポートで失われるため社名リスト自体は使わず、「その他（…）」の記入だけを回答として扱う。
CHECKLIST_NAMES = ["ウィルオブワーク", "エイジェック", "エヴァンド", "ネオキャリア", "ミラエール", "エムシャイン", "ブレクサネクスト"]
OTHER_FIELD = re.compile(r"その他\s*[（(](.*?)[）)]", re.S)
SKIP_TABS = ("運用ルール", "運用フロー", "学校訪問シート【原本】")


def clean(s: str) -> str:
    s = s.replace("\\*", "*").replace("**", "")
    s = re.sub(r"\\([~_&!#\[\]().\-+>|])", r"\1", s)
    s = s.replace("\u3000", " ")
    return s


def nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", s)


def split_tabs(text: str):
    parts = re.split(r"^# (.+?)\s*$", text, flags=re.M)
    for i in range(1, len(parts), 2):
        name = clean(parts[i]).strip().strip("【】").strip()
        yield name, parts[i + 1]


def table_cells(line: str):
    line = line.strip()
    if not (line.startswith("|") and line.endswith("|")):
        return None
    cells = [c.strip() for c in line[1:-1].split("|")]
    if all(re.fullmatch(r":?-+:?", c) for c in cells if c):
        return None  # 区切り行
    return cells


VISIT_ROW = re.compile(r"^\|\s*(?:\\\*\\\*)?訪問日")


def split_blocks(tab_text: str):
    lines = tab_text.splitlines()
    starts = [i for i, l in enumerate(lines) if VISIT_ROW.match(l)]
    if not starts:
        return [lines]
    blocks = []
    for k, s in enumerate(starts):
        e = starts[k + 1] if k + 1 < len(starts) else len(lines)
        # 1ブロック目の前にあるメモ（見出し直下の自由記述）は1ブロック目に含める
        b = lines[(0 if k == 0 else s):e]
        blocks.append(b)
    return blocks


DATE_RE = re.compile(r"(20\d{2})\s*[/年.]\s*(\d{1,2})\s*[/月.]\s*(\d{1,2})")


def parse_block(lines):
    blk = {"date": None, "school_field": None, "size_raw": None, "competitor": {}, "memo": []}
    size_next = False
    for raw in lines:
        line = clean(raw)
        cells = table_cells(line)
        if cells is not None:
            cells = [c for c in cells]
            label = nfkc(cells[0]) if cells else ""
            if label.startswith("訪問日") and len(cells) > 1 and blk["date"] is None:
                m = DATE_RE.search(nfkc(cells[1]))
                if m and m.group(2) != "0" and m.group(2) != "00":
                    blk["date"] = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                continue
            if label.startswith("学校名") and len(cells) > 1 and blk["school_field"] is None:
                blk["school_field"] = " ".join(c for c in cells[1:] if c.strip()).strip() or None
                continue
            if label.startswith("最上級生の全体人数"):
                size_next = True
                continue
            if size_next:
                size_next = False
                v = cells[0].strip() if cells else ""
                if v and nfkc(v) not in LABEL_WORDS:
                    blk["size_raw"] = v
                # 同じ行の「進路未定」等の欄にも人数が書かれることがあるのでメモにも残す
                rest = " / ".join(c for c in cells[1:] if c.strip())
                if rest:
                    blk["memo"].append(rest)
                continue
            hit = next((L for L in COMPETITOR_LABELS if L in nfkc(cells[0]) or L in cells[0]), None)
            if hit and len(cells) > 1:
                v = " ".join(c for c in cells[1:] if c.strip() and not nfkc(c).lstrip().startswith("例:")).strip()
                for p in TEMPLATE_PLACEHOLDERS:
                    v = v.replace(p, "").strip()
                if v:
                    blk["competitor"][hit] = v
                continue
            joined = " ".join(cells)
            if sum(n in joined for n in CHECKLIST_NAMES) >= 3:
                m = OTHER_FIELD.search(nfkc(joined))
                other = re.sub(r"[\s　]+", " ", m.group(1)).strip(" -&#10;") if m else ""
                if other:
                    blk["competitor"]["よく来ている競合他社（その他欄）"] = other
                continue
            # 記入例（「例：…」で始まる定型文）は除外する
            cells = [c for c in cells if not nfkc(c).lstrip().startswith("例:")]
            if label.startswith("★") and len(cells) > 1:
                v = " ".join(c for c in cells[1:] if c.strip()).strip()
                if v:
                    blk["competitor"][label.strip("★ ")[:30]] = v
                continue
            txt = " / ".join(c for c in cells if c.strip())
            if txt:
                blk["memo"].append(txt)
        else:
            t = line.strip()
            if t and not t.startswith("http"):
                blk["memo"].append(t)
    blk["memo"] = "\n".join(blk["memo"])
    return blk


def extract(doc_id: str, path: str):
    text = open(path, encoding="utf-8").read()
    out = []
    for name, body in split_tabs(text):
        if any(name.startswith(s) for s in SKIP_TABS):
            continue
        blocks = [parse_block(b) for b in split_blocks(body)]
        fields = [b["school_field"] for b in blocks if b["school_field"]]
        if "のコピー" in name:
            # 原本を複製したまま名前を変えていないタブは、シート内の「学校名」欄を学校名として使う
            if not fields:
                continue
            name = fields[0]
        out.append({"doc_id": doc_id, "tab": name, "school_fields": sorted(set(fields)), "blocks": blocks})
    return out


def main():
    out_path = sys.argv[1]
    result = []
    for arg in sys.argv[2:]:
        doc_id, path = arg.split("=", 1)
        result.extend(extract(doc_id, path))
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    print(f"{len(result)} tabs -> {out_path}")


if __name__ == "__main__":
    main()
