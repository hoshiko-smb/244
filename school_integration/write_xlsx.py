"""統合結果を管理表 xlsx の「27卒_アポイント＆訪問管理」タブにだけ書き込む。

    python3 write_xlsx.py 管理表.xlsx integration.json OUT.xlsx

- 他のタブ・スタイル・共有文字列などはバイト単位でそのままコピーする（openpyxl で保存し直すと
  画像・フィルタービュー・入力規則の一部が落ちるため、シートXMLだけを直接編集する）。
- 既存の列（A〜AM）は一切変更せず、空いている AN 列以降に追記する。
- 値は inlineStr で書くので sharedStrings.xml も変更しない。
"""
import json
import re
import sys
import zipfile
from xml.sax.saxutils import escape

SHEET_NAME = "27卒_アポイント＆訪問管理"

# (列, 見出し, 統合結果のキー, 列幅, 本文スタイル)
#   スタイル番号は同じタブで既に使われている書式を流用（270=見出し, 232=左寄せ枠線, 52=中央寄せ枠線）
COLUMNS = [
    ("AN", "学校規模\n（1学年）", "学校規模", 24, 232),
    ("AO", "AGIC\n訪問", "AGIC訪問", 8, 52),
    ("AP", "他社\n訪問", "他社訪問", 8, 52),
    ("AQ", "訪問競合社名\n（AGIC以外）", "訪問競合社名", 26, 232),
    ("AR", "競合情報\n時点", "情報時点", 11, 52),
    ("AS", "根拠：学校規模\n（訪問シート）", "根拠（学校規模）", 40, 232),
    ("AT", "根拠：競合\n（訪問シート・前回調査）", "根拠（競合）", 60, 232),
    ("AU", "抽出メモ\n（要確認理由）", "抽出メモ", 28, 232),
    ("AV", "担当確認\n（未確認/OK/要修正）", None, 12, 52),
    ("AW", "修正内容\n（担当記入）", None, 30, 232),
]
HEADER_STYLE = 270
CHECK_VALUES = '"未確認,OK,要修正"'


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


FIRST_NEW = col_to_num(COLUMNS[0][0])
LAST_NEW = col_to_num(COLUMNS[-1][0])


def cell_xml(ref: str, value: str | None, style: int) -> str:
    if value is None or value == "":
        return f'<c r="{ref}" s="{style}"/>'
    v = escape(str(value)).replace("\n", "&#10;")
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{v}</t></is></c>'


def sheet_path(z: zipfile.ZipFile) -> str:
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'<sheet [^>]*name="' + re.escape(SHEET_NAME) + r'"[^>]*r:id="(rId\d+)"', wb)
    rid = m.group(1)
    for rel in re.findall(r"<Relationship [^>]*/>", rels):
        if f'Id="{rid}"' in rel:
            return "xl/" + re.search(r'Target="([^"]+)"', rel).group(1).lstrip("/").removeprefix("xl/")
    raise SystemExit("sheet not found")


def patch_sheet(xml: str, records: dict[int, dict]) -> str:
    def patch_row(m):
        row_open, body = m.group(1), m.group(2)
        r = int(re.search(r'r="(\d+)"', row_open).group(1))
        # 既存セルのうち、新しく使う列（AN〜AW）にある「空の書式だけのセル」は置き換える
        kept = []
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', body, flags=re.S):
            n = col_to_num(cm.group(1))
            if FIRST_NEW <= n <= LAST_NEW:
                if "<v>" in cm.group(0) or "<is>" in cm.group(0) or "<f>" in cm.group(0):
                    raise SystemExit(f"row {r}: column {cm.group(1)} already has a value; refusing to overwrite")
                continue
            kept.append((n, cm.group(0)))
        new = []
        if r == 1:
            new = [cell_xml(f"{c}1", h, HEADER_STYLE) for c, h, _, _, _ in COLUMNS]
        elif r in records:
            rec = records[r]
            for c, _, key, _, style in COLUMNS:
                if key:
                    val = rec.get(key, "")
                elif c == "AV":
                    val = "未確認"
                else:
                    val = ""
                new.append(cell_xml(f"{c}{r}", val, style))
        before = [x for n, x in kept if n < FIRST_NEW]
        after = [x for n, x in kept if n > LAST_NEW]
        return row_open + "".join(before + new + after) + "</row>"

    xml = re.sub(r'(<row [^>]*>)(.*?)</row>', patch_row, xml, flags=re.S)

    # 列幅
    cols = "".join(
        f'<col customWidth="1" min="{col_to_num(c)}" max="{col_to_num(c)}" width="{w}"/>'
        for c, _, _, w, _ in COLUMNS
    )
    existing = re.search(r"<cols>(.*?)</cols>", xml, flags=re.S).group(1)
    for mx in re.findall(r'max="(\d+)"', existing):
        if int(mx) >= FIRST_NEW:
            raise SystemExit("existing <col> definitions overlap the new columns")
    xml = xml.replace("</cols>", cols + "</cols>", 1)

    # 使用範囲（dimension 要素がある場合のみ）
    xml = re.sub(r'<dimension ref="A1:[A-Z]+(\d+)"/>', lambda m: f'<dimension ref="A1:AW{m.group(1)}"/>', xml, count=1)

    # 担当確認列のプルダウン
    last_row = max(records) if records else 2
    dv = (f'<dataValidation type="list" allowBlank="1" showDropDown="0" showErrorMessage="1" '
          f'sqref="AV2:AV{last_row}"><formula1>{escape(CHECK_VALUES)}</formula1></dataValidation>')
    m = re.search(r"<dataValidations( count=\"(\d+)\")?>", xml)
    if not m:
        raise SystemExit("dataValidations element not found")
    if m.group(2):
        xml = xml.replace(m.group(0), f'<dataValidations count="{int(m.group(2)) + 1}">', 1)
    xml = xml.replace("</dataValidations>", dv + "</dataValidations>", 1)

    # 既存の基本フィルタ（シート上のフィルタ）の範囲を新しい列まで広げる（条件は変更しない）
    xml = re.sub(r'(<sheetData>.*</sheetData><autoFilter ref="\$A\$1:\$)([A-Z]+)(\$\d+")',
                 lambda m: m.group(1) + "AW" + m.group(3), xml, count=1, flags=re.S)
    return xml


def main():
    src, integ, out = sys.argv[1:4]
    records = {r["行"]: r for r in json.load(open(integ, encoding="utf-8"))}
    with zipfile.ZipFile(src) as zin:
        target = sheet_path(zin)
        with zipfile.ZipFile(out, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == target:
                    data = patch_sheet(data.decode("utf-8"), records).encode("utf-8")
                zout.writestr(info, data, compress_type=info.compress_type)
    print(f"patched {target} ({len(records)} rows) -> {out}")


if __name__ == "__main__":
    main()
