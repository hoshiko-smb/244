/**
 * 27卒_アポイント＆訪問管理 ― 学校規模 × 競合情報の取り込みと、担当者別の確認用フィルタービュー作成
 *
 * 使い方（本番の「アポイント＆訪問管理表」スプレッドシートで）
 *   1. 拡張機能 → Apps Script を開き、このファイルの内容を貼り付ける
 *   2. 左の「サービス」→「Google Sheets API」を追加（フィルタービューの作成に必要）
 *   3. 統合済み xlsx を Google ドライブにアップロードして Google スプレッドシートとして開き、
 *      その URL の /d/ と /edit の間の ID を SOURCE_SPREADSHEET_ID に貼る
 *   4. importIntegration → setupCheckColumn → createReviewFilterViews の順に実行
 *
 * 書き込むのは「27卒_アポイント＆訪問管理」タブの AO〜AX 列（追加列）だけ。既存の A〜AN 列には触れない。
 */

const SHEET_NAME = '27卒_アポイント＆訪問管理';
const SOURCE_SPREADSHEET_ID = 'ここに統合済みファイルのIDを貼る';

const COL_SCHOOL = 8;    // H 学校名
const COL_VISITOR = 20;  // T 訪問者
// 本体の配置：AN=学校名（照合用）、AO〜AX=追加データ（2026/11/9 時点）
const COL_FIRST = 41;    // AO 学校規模（1学年）
const COL_AGIC = 42;     // AP AGIC訪問
const COL_CHECK = 49;    // AW 担当確認
const COL_LAST = 50;     // AX 修正内容
const NUM_NEW = COL_LAST - COL_FIRST + 1;
const SRC_FIRST = 40;    // 統合済み xlsx 側の先頭列（AN 学校規模）

/**
 * 統合済みファイルの AN〜AW を本番シートの AO〜AX へ写す（本番は AN を学校名の照合列にしている）。
 * 行番号だけでなく学校名（H列）も一致した行にだけ書くので、エクスポート後に行の並べ替え・追加が
 * あっても誤った学校に書き込まない。ずれていた行はログに出す。
 */
function importIntegration() {
  const src = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const dst = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const srcRows = src.getLastRow();
  const srcSchool = src.getRange(1, COL_SCHOOL, srcRows, 1).getValues();
  const srcNew = src.getRange(1, SRC_FIRST, srcRows, NUM_NEW).getValues();

  // 本番シートの学校名 → 行番号（同名の学校が複数ある場合は全行）
  const dstRows = dst.getLastRow();
  const dstSchool = dst.getRange(1, COL_SCHOOL, dstRows, 1).getValues();
  const dstCheck = dst.getRange(1, COL_CHECK, dstRows, 2).getValues();  // 既存の AV/AW（担当者の記入）
  const index = {};
  dstSchool.forEach((r, i) => {
    const k = String(r[0]).trim();
    if (k) (index[k] = index[k] || []).push(i + 1);
  });

  // 見出し
  dst.getRange(1, COL_FIRST, 1, NUM_NEW).setValues([srcNew[0]]);

  const unmatched = [];
  let written = 0;
  for (let i = 1; i < srcRows; i++) {
    const school = String(srcSchool[i][0]).trim();
    if (!school) continue;
    const rowNo = i + 1;
    let target = null;
    if (String(dstSchool[rowNo - 1] ? dstSchool[rowNo - 1][0] : '').trim() === school) {
      target = rowNo;                                   // 同じ行に同じ学校がいる
    } else if (index[school] && index[school].length === 1) {
      target = index[school][0];                        // 行が動いていたが学校名で一意に特定できる
    }
    if (!target) { unmatched.push(`${rowNo}行目 ${school}`); continue; }
    // 既に担当者が確認・修正を入れている場合は AV/AW を上書きしない
    const current = dstCheck[target - 1];
    const row = srcNew[i].slice();
    if (current[0] && current[0] !== '未確認') { row[COL_CHECK - COL_FIRST] = current[0]; }
    if (current[1]) { row[COL_LAST - COL_FIRST] = current[1]; }
    dst.getRange(target, COL_FIRST, 1, NUM_NEW).setValues([row]);
    written++;
  }
  Logger.log(`書き込み ${written} 行 / 照合できず ${unmatched.length} 行`);
  if (unmatched.length) Logger.log('照合できなかった行:\n' + unmatched.join('\n'));
}

/** 担当確認（AV列）のプルダウンを設定する。 */
function setupCheckColumn() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未確認', 'OK', '要修正'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, COL_CHECK, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

/**
 * 訪問者ごとの「確認用」フィルタービューと、戦略検討用のビューを作る（同名のビューは作り直す）。
 *  - 【確認用】◯◯   : 訪問者＝◯◯ かつ 担当確認が OK 以外
 *  - 【戦略】AGIC訪問あり / 【戦略】AGIC不明（要ヒアリング）
 * フィルタービューは各自の表示だけが変わり、他の人の画面や元データには影響しない。
 */
function createReviewFilterViews() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  const sheetId = sh.getSheetId();
  const lastRow = sh.getLastRow();
  const range = { sheetId, startRowIndex: 0, endRowIndex: lastRow, startColumnIndex: 0, endColumnIndex: COL_LAST };

  const visitors = [...new Set(
    sh.getRange(2, COL_VISITOR, lastRow - 1, 1).getValues().map(r => String(r[0]).trim()).filter(v => v)
  )].sort();

  const views = visitors.map(name => ({
    title: `【確認用】${name}`,
    range,
    filterSpecs: [
      { columnIndex: COL_VISITOR - 1, filterCriteria: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: name }] } } },
      { columnIndex: COL_CHECK - 1, filterCriteria: { hiddenValues: ['OK'] } },
    ],
    sortSpecs: [{ dimensionIndex: COL_SCHOOL - 1, sortOrder: 'ASCENDING' }],
  }));
  views.push({
    title: '【戦略】AGIC訪問あり',
    range,
    filterSpecs: [{ columnIndex: COL_AGIC - 1, filterCriteria: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'あり' }] } } }],
  });
  views.push({
    title: '【戦略】AGIC不明（要ヒアリング）',
    range,
    filterSpecs: [{ columnIndex: COL_AGIC - 1, filterCriteria: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: '不明' }] } } }],
  });

  // 同じタイトルの既存ビューは削除してから作り直す（範囲・条件を最新にするため）
  const meta = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(properties(sheetId),filterViews(filterViewId,title))' });
  const existing = (meta.sheets.find(s => s.properties.sheetId === sheetId).filterViews || []);
  const titles = new Set(views.map(v => v.title));
  const requests = existing
    .filter(v => titles.has(v.title))
    .map(v => ({ deleteFilterView: { filterId: v.filterViewId } }));
  views.forEach(v => requests.push({ addFilterView: { filter: v } }));
  Sheets.Spreadsheets.batchUpdate({ requests }, ss.getId());
  Logger.log(`フィルタービュー ${views.length} 件を作成: ${views.map(v => v.title).join(', ')}`);
}
