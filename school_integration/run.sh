#!/usr/bin/env bash
# 学校規模 × 競合情報の統合を一括実行する。
#
#   ./run.sh 管理表.xlsx 訪問シートのテキストを置いたフォルダ 出力フォルダ [締め日 YYYY-MM-DD]
#
# 訪問シートのフォルダには、Google ドキュメントを Drive の read_file_content 等で取得した本文を
# 「<ドキュメントID>.txt」という名前で置く（例: 1W5ZdoK65C9on3RhqbWY5WOq6wZJ8qaELjMVL9LKmYoA.txt）。
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
XLSX="$1"; DOCS="$2"; OUT="$3"; CUTOFF="${4:-2026-09-30}"
mkdir -p "$OUT"

args=()
for f in "$DOCS"/*.txt; do
  args+=("$(basename "$f" .txt)=$f")
done

python3 "$HERE/extract_visit_sheets.py" "$OUT/tabs.json" "${args[@]}"
python3 "$HERE/match_rows.py" "$XLSX" "$OUT/tabs.json" "$OUT/match.json" "$HERE/manual_matches.json"
PYTHONPATH="$HERE" python3 "$HERE/build_integration.py" "$XLSX" "$OUT/tabs.json" "$OUT/match.json" \
  "$HERE/size_overrides.json" "$OUT" --cutoff "$CUTOFF"
python3 "$HERE/summarize.py" "$XLSX" "$OUT/integration.json" "$OUT" > /dev/null
python3 "$HERE/write_xlsx.py" "$XLSX" "$OUT/integration.json" "$OUT/管理表_27卒_学校規模・競合統合.xlsx"
python3 "$HERE/make_paste_sheet.py" "$XLSX" "$OUT/integration.json" "$OUT/コピペ用_本体AN1に貼り付け.xlsx"
echo "done: $OUT"
