"""Google スプレッドシートの本体へ「コピペ」で反映するためのファイルを作る。

    python3 make_paste_sheet.py 管理表.xlsx integration.json OUT.xlsx [--order-sheet シート26]

--order-sheet を付けると、そのシート（A列=No.、B列=学校名）の並び順で出力する。
管理表タブのA列（No.）と対応付けるので、同名の学校が複数あってもずれない。

- 「貼り付け用」シートの行番号は、管理表「27卒_アポイント＆訪問管理」タブの行番号と完全に一致させる
  （1行目=見出し、2行目以降=同じ行の学校）。A〜J列を選んでコピーし、本体の AN1 に貼れば反映できる。
- 見出しは1行（改行なし）にして、テキスト経由で貼っても崩れないようにする。
- L列に「照合用」の学校名を置く（コピー対象外）。貼る前に本体のH列と並びが同じか確認できる。
"""
import json
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

SHEET = "27卒_アポイント＆訪問管理"
# 本体での列, 見出し, 統合結果のキー, 列幅
COLUMNS = [
    ("AN", "学校規模（1学年）", "学校規模", 30),
    ("AO", "AGIC訪問", "AGIC訪問", 10),
    ("AP", "他社訪問", "他社訪問", 10),
    ("AQ", "訪問競合社名（AGIC以外）", "訪問競合社名", 28),
    ("AR", "競合情報時点", "情報時点", 12),
    ("AS", "根拠：学校規模", "根拠（学校規模）", 40),
    ("AT", "根拠：競合", "根拠（競合）", 50),
    ("AU", "抽出メモ（要確認理由）", "抽出メモ", 28),
    ("AV", "担当確認（未確認/OK/要修正）", None, 14),
    ("AW", "修正内容（担当記入）", None, 24),
]
HEAD_FILL = PatternFill("solid", fgColor="FFF2CC")
CHECK_FILL = PatternFill("solid", fgColor="EDEDED")


def load_order(wb, order_sheet):
    """並び順シートの行ごとに、管理表タブの行番号を返す。{出力行: (管理表の行, No., 学校名)}"""
    ws_src = wb[SHEET]
    by_no = {}
    for i, row in enumerate(ws_src.iter_rows(values_only=True), start=1):
        if i > 1 and row[7] and row[0] is not None:
            by_no[int(row[0])] = (i, row[7])
    order = {}
    for i, row in enumerate(wb[order_sheet].iter_rows(values_only=True), start=1):
        if i == 1 or row[0] is None:
            continue
        no, name = int(row[0]), row[1]
        if no not in by_no:
            raise SystemExit(f"{order_sheet} {i}行目 No.{no} が管理表のA列にありません")
        src_row, src_name = by_no[no]
        if str(src_name).strip() != str(name).strip():
            raise SystemExit(f"No.{no}: {order_sheet}「{name}」と管理表「{src_name}」が一致しません")
        order[i] = (src_row, no, name)
    return order


def main():
    args = sys.argv[1:]
    order_sheet = None
    if "--order-sheet" in args:
        k = args.index("--order-sheet")
        order_sheet = args[k + 1]
        del args[k:k + 2]
    src, integ, out = args[:3]
    recs = {r["行"]: r for r in json.load(open(integ, encoding="utf-8"))}
    wb_src = openpyxl.load_workbook(src, read_only=True)
    ws_src = wb_src[SHEET]
    schools = {i: row[7] for i, row in enumerate(ws_src.iter_rows(values_only=True), start=1)}
    if order_sheet:
        order = load_order(wb_src, order_sheet)          # 出力行 -> (管理表の行, No., 学校名)
    else:
        order = {r: (r, None, s) for r, s in schools.items() if r > 1 and s}
    last = max(order)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"貼り付け用（{order_sheet}の順）" if order_sheet else "貼り付け用（本体AN1へ）"
    for j, (col, head, _, width) in enumerate(COLUMNS, start=1):
        c = ws.cell(1, j, head)
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = width
    ws.cell(1, 12, "照合用：学校名 ※コピー不要").font = Font(bold=True)
    ws.cell(1, 13, f"照合用：{order_sheet}のNo." if order_sheet else "照合用：本体の行番号").font = Font(bold=True)
    for j in (12, 13):
        ws.cell(1, j).fill = CHECK_FILL
    ws.column_dimensions["K"].width = 3
    ws.column_dimensions["L"].width = 40
    ws.column_dimensions["M"].width = 12

    for r in range(2, last + 1):
        src_row, no, name = order.get(r, (None, None, None))
        rec = recs.get(src_row)
        if rec:
            for j, (col, _, key, _) in enumerate(COLUMNS, start=1):
                v = rec.get(key, "") if key else ("未確認" if col == "AV" else "")
                if v:
                    ws.cell(r, j, v)
        if name:
            ws.cell(r, 12, name).fill = CHECK_FILL
        ws.cell(r, 13, no if order_sheet else r).fill = CHECK_FILL
    dv = DataValidation(type="list", formula1='"未確認,OK,要修正"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"I2:I{last}")
    ws.freeze_panes = "A2"

    guide = wb.create_sheet("使い方")
    if order_sheet:
        first, final = order[min(order)][2], order[last][2]
        lines = [
            f"■ この「貼り付け用」シートは「{order_sheet}」の並び順（No.1〜{order[last][1]}）に合わせています",
            f"  2行目＝No.1（{first}）… {last}行目＝No.{order[last][1]}（{final}）。{order_sheet} の行番号と1対1で対応します。",
            "",
            "■ 貼り付け手順",
            f"1. L列（学校名）・M列（No.）が、貼り付け先の学校名・No.と同じ並びか確認する",
            "2. 貼り付け先でフィルタがかかっている場合は解除して全行を表示する（非表示の行があると位置がずれるため）",
            f"3. 「貼り付け用」シートの A1:J{last} を選択してコピー（L・M列はコピー不要）",
            f"4. 貼り付け先の見出し行・先頭列のセルを選択し、値のみ貼り付け（Ctrl+Shift+V／Mac は ⌘+Shift+V）",
            f"   例：{order_sheet} の C1 に貼る／{SHEET} タブを A列（No.）の昇順に並べ替えてから AN1 に貼る",
            "5. 担当確認列のプルダウンが必要な場合は、貼り付け後に「データ」→「データの入力規則」で「未確認,OK,要修正」を設定",
            "",
            "■ 列の意味",
            "学校規模（1学年○○人程度） / AGIC（エイジェック）訪問 あり・なし・不明 / AGIC以外の競合の訪問 / 競合の社名",
            "競合情報の時点 / 根拠（訪問シートの記述の抜粋） / 要確認の理由 / 担当確認 / 担当者の修正内容",
            "競合情報は 2026-09-30 までの訪問分で確定。学校規模は最新の訪問シートの記載。",
        ]
    else:
        lines = [
        "■ Google スプレッドシート（本体）への反映手順",
        "1. このファイルを Google ドライブにアップロードし、Google スプレッドシートで開く（Excel で開いても可）",
        f"2. 「貼り付け用」シートの L列（学校名）と、本体「{SHEET}」タブの H列が同じ並びか確認する",
        f"   （例: 2行目={schools.get(2)}、{last}行目={schools.get(last)}。行の追加・並べ替えをしていなければ一致します）",
        "3. 本体タブでフィルタがかかっている場合は、先にフィルタを解除して全行を表示する（非表示の行があると貼り付け位置がずれるため）",
        f"4. 「貼り付け用」シートの A1:J{last} を選択してコピー",
        f"5. 本体「{SHEET}」タブの AN1 セルを選択し、値のみ貼り付け（Ctrl+Shift+V／Mac は ⌘+Shift+V）",
        "6. AV列（担当確認）にプルダウンが必要な場合は、AV列を選択 →「データ」→「データの入力規則」で「未確認,OK,要修正」を設定",
        "",
        "■ 列の意味",
        "AN 学校規模（1学年○○人程度） / AO AGIC（エイジェック）訪問 あり・なし・不明 / AP AGIC以外の競合の訪問 / AQ 競合の社名",
        "AR 競合情報の時点 / AS・AT 根拠（訪問シートの記述の抜粋） / AU 要確認の理由 / AV 担当確認 / AW 担当者の修正内容",
        "競合情報は 2026-09-30 までの訪問分で確定。学校規模は最新の訪問シートの記載。",
    ]
    for i, t in enumerate(lines, start=1):
        guide.cell(i, 1, t)
    guide.column_dimensions["A"].width = 120
    wb.save(out)
    print(f"rows 2..{last} -> {out}")


if __name__ == "__main__":
    main()
